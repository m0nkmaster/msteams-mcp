/**
 * Calendar API client for meeting operations.
 * 
 * Handles calls to teams.microsoft.com/api/mt/part/{region}/v2.1/me/calendars endpoints.
 */

import { httpRequest } from '../utils/http.js';
import {
  CALENDAR_API,
  GRAPH_CALENDAR_API,
  getTeamsHeaders,
  getGraphHeaders,
} from '../utils/api-config.js';
import { type Result, ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import { requireSkypeSpacesAuth, requireGraphAuth, getRegionConfig } from '../utils/auth-guards.js';
import {
  DEFAULT_MEETING_LIMIT,
  DEFAULT_MEETING_DAYS_AHEAD,
} from '../constants.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/** Organiser information for a meeting. */
export interface MeetingOrganizer {
  name: string;
  email: string;
}

/** A calendar meeting/event. */
export interface Meeting {
  /** Unique identifier for the meeting. */
  id: string;
  /** Meeting subject/title. */
  subject: string;
  /** Start time (ISO 8601). */
  startTime: string;
  /** End time (ISO 8601). */
  endTime: string;
  /** Meeting organiser. */
  organizer: MeetingOrganizer;
  /** Location (room name or text). */
  location?: string;
  /** Whether this is a Teams online meeting. */
  isOnlineMeeting: boolean;
  /** Teams join URL (if online meeting). */
  joinUrl?: string;
  /** Meeting chat thread ID (for use with teams_get_thread). */
  threadId?: string;
  /** URL used to update or remove the Teams scheduling service entry for this meeting. */
  schedulingServiceUpdateUrl?: string;
  /** Your RSVP status. */
  myResponse: 'None' | 'Accepted' | 'Tentative' | 'Declined';
  /** Calendar show-as status. */
  showAs: 'Free' | 'Busy' | 'Tentative' | 'OutOfOffice' | 'Unknown';
  /** Whether you're the organiser. */
  isOrganizer: boolean;
  /** Event type (Single, Occurrence, Exception, SeriesMaster). */
  eventType: string;
}

/** Options for fetching calendar view. */
export interface CalendarViewOptions {
  /** Start of date range (ISO 8601). Defaults to now. */
  startDate?: string;
  /** End of date range (ISO 8601). Defaults to 7 days from now. */
  endDate?: string;
  /** Maximum results to return. */
  limit?: number;
}

/** Response from getting calendar view. */
export interface CalendarViewResult {
  /** Number of meetings returned. */
  count: number;
  /** List of meetings. */
  meetings: Meeting[];
}

/** An attendee for a meeting invite. */
export interface MeetingAttendee {
  /** Email address of the attendee. */
  email: string;
  /** Display name of the attendee (optional). */
  name?: string;
}

/** Options for creating a new calendar event. */
export interface CreateMeetingOptions {
  /** Meeting subject/title. */
  subject: string;
  /** Start time (ISO 8601, e.g. "2026-06-05T10:00:00Z"). */
  startTime: string;
  /** End time (ISO 8601, e.g. "2026-06-05T10:30:00Z"). */
  endTime: string;
  /** Attendees to invite (email addresses). */
  attendees?: MeetingAttendee[];
  /** Meeting body/description text. */
  body?: string;
  /** Whether to create as a Teams online meeting (default: true). */
  isOnlineMeeting?: boolean;
  /** Physical or virtual location. */
  location?: string;
}

/** Result from creating a calendar event. */
export interface CreatedMeeting {
  /** The newly created meeting's ID. */
  id: string;
  /** Meeting subject. */
  subject: string;
  /** Start time (ISO 8601). */
  startTime: string;
  /** End time (ISO 8601). */
  endTime: string;
  /** Teams join URL (if online meeting). */
  joinUrl?: string;
}

/** Options for updating an existing calendar event. All fields are optional. */
export type UpdateMeetingOptions = Partial<Omit<CreateMeetingOptions, 'isOnlineMeeting'>>;

/** RSVP response type. */
export type MeetingResponse = 'accept' | 'tentativelyAccept' | 'decline';

/** Options for responding to a meeting invite. */
export interface RespondToMeetingOptions {
  /** Your response. */
  response: MeetingResponse;
  /** Optional message to include with your response. */
  comment?: string;
  /**
   * Whether to send a response email to the organiser (default: true).
   * Set to false to silently update your calendar without notifying the organiser.
   */
  sendResponse?: boolean;
  /**
   * For decline only — propose an alternative time.
   * Only valid when response is "decline" and the event allows new time proposals.
   */
  proposedNewTime?: {
    start: string; // ISO 8601
    end: string;   // ISO 8601
  };
}

/** A single free/busy slot for a user. */
export interface ScheduleItem {
  /** Start of the busy slot (ISO 8601). */
  start: string;
  /** End of the busy slot (ISO 8601). */
  end: string;
  /** Free/busy status for this slot. */
  status: 'free' | 'busy' | 'tentative' | 'oof' | 'workingElsewhere' | 'unknown';
  /** Subject of the event (may be empty for privacy). */
  subject?: string;
  /** Whether the event is private. */
  isPrivate?: boolean;
}

/** Free/busy schedule for a single user. */
export interface UserSchedule {
  /** The email address / identity this schedule belongs to. */
  scheduleId: string;
  /**
   * String encoding of availability in the requested window.
   * Each character represents one interval (default 30 min):
   * 0=free, 1=tentative, 2=busy, 3=OOF, 4=working elsewhere
   */
  availabilityView: string;
  /** Individual busy slots within the requested window. */
  scheduleItems: ScheduleItem[];
}

/** Options for fetching free/busy schedules. */
export interface GetScheduleOptions {
  /** Email addresses of people whose availability to check. */
  schedules: string[];
  /** Start of the window to check (ISO 8601). */
  startTime: string;
  /** End of the window to check (ISO 8601). */
  endTime: string;
  /**
   * Duration of each slot in the availabilityView string, in minutes.
   * Default: 30. Min: 5. Max: 1440.
   */
  availabilityViewInterval?: number;
}

/** Result from getSchedule. */
export interface GetScheduleResult {
  schedules: UserSchedule[];
}

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Graph `Prefer` header that asks for all event times to be returned in UTC.
 * Combined with {@link toUtcIso}, this guarantees unambiguous ISO 8601 output.
 */
const GRAPH_UTC_PREFER = 'outlook.timezone="UTC"';

/**
 * Normalises a Graph `dateTimeTimeZone` value to an unambiguous ISO 8601 string.
 *
 * Graph returns `dateTime` without any timezone marker (e.g.
 * "2026-06-05T10:00:00.0000000") alongside a separate `timeZone` field. When the
 * times are UTC (which we force via the `Prefer: outlook.timezone="UTC"` header),
 * we append a `Z` so consumers don't misread them as local time.
 */
export function toUtcIso(dt?: { dateTime?: string; timeZone?: string }): string {
  const raw = dt?.dateTime;
  if (!raw) return '';
  // Already carries an offset or Z.
  if (/([zZ]|[+-]\d{2}:?\d{2})$/.test(raw)) return raw;
  // Mark UTC times as such; leave non-UTC values untouched.
  if (!dt?.timeZone || dt.timeZone.toUpperCase() === 'UTC') return `${raw}Z`;
  return raw;
}

/**
 * Parses a raw API meeting response into our Meeting type.
 */
function parseMeeting(raw: Record<string, unknown>): Meeting {
  // Extract thread ID from skypeTeamsData if present
  let threadId: string | undefined;
  const skypeTeamsData = raw.skypeTeamsData as string | undefined;
  if (skypeTeamsData) {
    try {
      const parsed = JSON.parse(skypeTeamsData) as Record<string, unknown>;
      threadId = parsed.cid as string | undefined;
    } catch {
      // Ignore parsing errors
    }
  }

  // Map response type to our enum values
  const rawResponse = raw.myResponseType as string | undefined;
  let myResponse: Meeting['myResponse'] = 'None';
  if (rawResponse === 'Accepted' || rawResponse === 'Organizer') {
    myResponse = 'Accepted';
  } else if (rawResponse === 'Tentative' || rawResponse === 'TentativelyAccepted') {
    myResponse = 'Tentative';
  } else if (rawResponse === 'Declined') {
    myResponse = 'Declined';
  }

  // Map showAs values
  const rawShowAs = raw.showAs as string | undefined;
  let showAs: Meeting['showAs'] = 'Unknown';
  if (rawShowAs === 'Free') {
    showAs = 'Free';
  } else if (rawShowAs === 'Busy') {
    showAs = 'Busy';
  } else if (rawShowAs === 'Tentative') {
    showAs = 'Tentative';
  } else if (rawShowAs === 'Oof' || rawShowAs === 'OutOfOffice') {
    showAs = 'OutOfOffice';
  }

  return {
    id: raw.objectId as string,
    subject: (raw.subject as string) || '(No subject)',
    startTime: raw.startTime as string,
    endTime: raw.endTime as string,
    organizer: {
      name: (raw.organizerName as string) || 'Unknown',
      email: (raw.organizerAddress as string) || '',
    },
    location: raw.location as string | undefined,
    isOnlineMeeting: raw.isOnlineMeeting === true,
    joinUrl: raw.skypeTeamsMeetingUrl as string | undefined,
    threadId,
    schedulingServiceUpdateUrl: raw.schedulingServiceUpdateUrl as string | undefined,
    myResponse,
    showAs,
    isOrganizer: raw.isOrganizer === true,
    eventType: (raw.eventType as string) || 'Single',
  };
}

/**
 * Parses a raw Microsoft Graph event into our Meeting type.
 *
 * Graph returns a different shape than the mt/part calendarView, so this maps
 * the Graph fields onto the same Meeting interface for consistency.
 */
function parseGraphEvent(raw: Record<string, unknown>): Meeting {
  const organizer = raw.organizer as { emailAddress?: { name?: string; address?: string } } | undefined;
  const location = raw.location as { displayName?: string } | undefined;
  const onlineMeeting = raw.onlineMeeting as { joinUrl?: string } | undefined;
  const start = raw.start as { dateTime?: string; timeZone?: string } | undefined;
  const end = raw.end as { dateTime?: string; timeZone?: string } | undefined;
  const responseStatus = raw.responseStatus as { response?: string } | undefined;

  // Map Graph responseStatus to our enum
  let myResponse: Meeting['myResponse'] = 'None';
  switch (responseStatus?.response) {
    case 'organizer':
    case 'accepted':
      myResponse = 'Accepted';
      break;
    case 'tentativelyAccepted':
      myResponse = 'Tentative';
      break;
    case 'declined':
      myResponse = 'Declined';
      break;
  }

  // Map Graph showAs to our enum
  const rawShowAs = raw.showAs as string | undefined;
  let showAs: Meeting['showAs'] = 'Unknown';
  if (rawShowAs === 'free') showAs = 'Free';
  else if (rawShowAs === 'busy') showAs = 'Busy';
  else if (rawShowAs === 'tentative') showAs = 'Tentative';
  else if (rawShowAs === 'oof' || rawShowAs === 'workingElsewhere') showAs = 'OutOfOffice';

  return {
    id: raw.id as string,
    subject: (raw.subject as string) || '(No subject)',
    startTime: toUtcIso(start),
    endTime: toUtcIso(end),
    organizer: {
      name: organizer?.emailAddress?.name ?? 'Unknown',
      email: organizer?.emailAddress?.address ?? '',
    },
    location: location?.displayName,
    isOnlineMeeting: raw.isOnlineMeeting === true,
    joinUrl: onlineMeeting?.joinUrl,
    threadId: undefined,
    schedulingServiceUpdateUrl: undefined,
    myResponse,
    showAs,
    isOrganizer: raw.isOrganizer === true,
    eventType: (raw.type as string) || 'singleInstance',
  };
}

/**
 * Builds a Graph event request body from create/update options.
 */
function buildGraphEventBody(options: Partial<CreateMeetingOptions>): Record<string, unknown> {
  const body: Record<string, unknown> = {};
  if (options.subject !== undefined) body.subject = options.subject;
  if (options.startTime !== undefined) body.start = { dateTime: options.startTime, timeZone: 'UTC' };
  if (options.endTime !== undefined) body.end = { dateTime: options.endTime, timeZone: 'UTC' };
  if (options.body !== undefined) body.body = { contentType: 'text', content: options.body };
  if (options.location !== undefined) body.location = { displayName: options.location };
  if (options.attendees !== undefined) {
    body.attendees = options.attendees.map((a) => ({
      emailAddress: { address: a.email, name: a.name ?? a.email },
      type: 'required',
    }));
  }
  if (options.isOnlineMeeting !== undefined) {
    body.isOnlineMeeting = options.isOnlineMeeting;
    if (options.isOnlineMeeting) body.onlineMeetingProvider = 'teamsForBusiness';
  }
  return body;
}

/**
 * Builds the select fields for the calendar API.
 */
function getSelectFields(): string {
  return [
    'cleanGlobalObjectId',
    'endTime',
    'eventTimeZone',
    'eventType',
    'hasAttachments',
    'iCalUid',
    'isAllDayEvent',
    'isAppointment',
    'isCancelled',
    'isOnlineMeeting',
    'isOrganizer',
    'isPrivate',
    'lastModifiedTime',
    'location',
    'myResponseType',
    'objectId',
    'organizerAddress',
    'organizerName',
    'schedulingServiceUpdateUrl',
    'showAs',
    'skypeTeamsData',
    'skypeTeamsMeetingUrl',
    'startTime',
    'subject',
  ].join(',');
}

// ─────────────────────────────────────────────────────────────────────────────
// API Functions
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Gets meetings from the user's calendar for a date range.
 * 
 * The region and partition are extracted from the user's session (DISCOVER-REGION-GTM),
 * so we always use the correct endpoint without guessing.
 * 
 * @param options - Options for filtering meetings
 * @returns List of meetings
 */
export async function getCalendarView(
  options: CalendarViewOptions = {}
): Promise<Result<CalendarViewResult>> {
  const authResult = requireSkypeSpacesAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const { skypeToken, spacesToken } = authResult.value;

  const regionConfig = getRegionConfig();
  if (!regionConfig) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'Could not determine region. Please run teams_login to authenticate.',
      { suggestions: ['Call teams_login to authenticate'] }
    ));
  }

  // Calculate default date range
  const now = new Date();
  const defaultEnd = new Date(now);
  defaultEnd.setDate(defaultEnd.getDate() + DEFAULT_MEETING_DAYS_AHEAD);

  const startDate = options.startDate || now.toISOString();
  const endDate = options.endDate || defaultEnd.toISOString();
  const limit = options.limit ?? DEFAULT_MEETING_LIMIT;

  // Build query parameters
  const params = new URLSearchParams({
    startDate,
    endDate,
    '$top': limit.toString(),
    '$count': 'true',
    '$skip': '0',
    '$orderby': 'startTime asc',
    // Filter out appointments (blocks), all-day events, and cancelled meetings
    '$filter': 'isAppointment eq false and isAllDayEvent eq false and isCancelled eq false',
    '$select': getSelectFields(),
  });

  // Use the exact region-partition from session discovery
  // Some tenants use partitioned URLs (/api/mt/part/amer-02), others don't (/api/mt/emea)
  const calendarUrl = CALENDAR_API.calendarView(regionConfig.regionPartition, regionConfig.hasPartition, regionConfig.teamsBaseUrl);
  const url = `${calendarUrl}?${params.toString()}`;

  const response = await httpRequest<Record<string, unknown>>(
    url,
    {
      method: 'GET',
      headers: {
        ...getTeamsHeaders(regionConfig.teamsBaseUrl),
        'Authentication': `skypetoken=${skypeToken}`,
        'Authorization': `Bearer ${spacesToken}`,
      },
    }
  );

  if (!response.ok) {
    return response;
  }

  const data = response.value.data;
  const rawMeetings = data.value as Array<Record<string, unknown>> | undefined;

  if (!rawMeetings || !Array.isArray(rawMeetings)) {
    return ok({
      count: 0,
      meetings: [],
    });
  }

  const meetings = rawMeetings.map(parseMeeting);

  return ok({
    count: meetings.length,
    meetings,
  });
}

/**
 * Creates a new calendar event (meeting) in the user's Teams calendar.
 *
 * Uses the Microsoft Graph API (POST /me/events) with a Graph token.
 * Optionally adds a Teams online meeting link and sends invites to attendees.
 *
 * @param options - Event creation options
 * @returns The created meeting's ID, subject, times, and join URL
 */
export async function createMeeting(
  options: CreateMeetingOptions
): Promise<Result<CreatedMeeting>> {
  const authResult = requireGraphAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const graphToken = authResult.value;

  const body = buildGraphEventBody({
    ...options,
    isOnlineMeeting: options.isOnlineMeeting !== false,
  });

  const response = await httpRequest<Record<string, unknown>>(
    GRAPH_CALENDAR_API.events(),
    {
      method: 'POST',
      headers: { ...getGraphHeaders(graphToken), 'Prefer': GRAPH_UTC_PREFER },
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    return response;
  }

  const data = response.value.data as Record<string, unknown>;
  const onlineMeeting = data.onlineMeeting as Record<string, unknown> | undefined;
  const joinUrl = onlineMeeting?.joinUrl as string | undefined;

  return ok({
    id: data.id as string,
    subject: (data.subject as string) ?? options.subject,
    startTime: toUtcIso(data.start as { dateTime?: string; timeZone?: string }) || options.startTime,
    endTime: toUtcIso(data.end as { dateTime?: string; timeZone?: string }) || options.endTime,
    joinUrl,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Get single meeting
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Gets a single calendar event by its ID.
 *
 * @param eventId - The event ID (meeting.id from getCalendarView or createMeeting)
 */
export async function getMeeting(eventId: string): Promise<Result<Meeting>> {
  const authResult = requireGraphAuth();
  if (!authResult.ok) return authResult;
  const graphToken = authResult.value;

  const response = await httpRequest<Record<string, unknown>>(
    GRAPH_CALENDAR_API.event(eventId),
    {
      method: 'GET',
      headers: { ...getGraphHeaders(graphToken), 'Prefer': GRAPH_UTC_PREFER },
    }
  );

  if (!response.ok) return response;

  const raw = response.value.data as Record<string, unknown>;
  return ok(parseGraphEvent(raw));
}

// ─────────────────────────────────────────────────────────────────────────────
// Update meeting
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Updates an existing calendar event.
 *
 * Only the fields you provide are changed (PATCH semantics).
 *
 * @param eventId - The event ID to update
 * @param options - Fields to change
 */
export async function updateMeeting(
  eventId: string,
  options: UpdateMeetingOptions
): Promise<Result<Meeting>> {
  const authResult = requireGraphAuth();
  if (!authResult.ok) return authResult;
  const graphToken = authResult.value;

  const body = buildGraphEventBody(options);

  const response = await httpRequest<Record<string, unknown>>(
    GRAPH_CALENDAR_API.event(eventId),
    {
      method: 'PATCH',
      headers: { ...getGraphHeaders(graphToken), 'Prefer': GRAPH_UTC_PREFER },
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) return response;

  const raw = response.value.data as Record<string, unknown>;
  return ok(parseGraphEvent(raw));
}

// ─────────────────────────────────────────────────────────────────────────────
// Cancel / delete meeting
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Cancels or removes a calendar event.
 *
 * If you are the organiser, this sends a cancellation to all attendees.
 * If you are an attendee, this removes the event from your calendar only.
 *
 * @param eventId - The event ID to cancel/remove
 */
export async function cancelMeeting(eventId: string): Promise<Result<void>> {
  const authResult = requireGraphAuth();
  if (!authResult.ok) return authResult;
  const graphToken = authResult.value;

  const response = await httpRequest<void>(
    GRAPH_CALENDAR_API.event(eventId),
    {
      method: 'DELETE',
      headers: getGraphHeaders(graphToken),
    }
  );

  if (!response.ok) return response;
  return ok(undefined);
}

// ─────────────────────────────────────────────────────────────────────────────
// RSVP — respond to a meeting invite
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Accepts, tentatively accepts, or declines a meeting invite.
 *
 * @param eventId - The event ID to respond to
 * @param options - Response type, optional comment, and optional proposed new time
 */
export async function respondToMeeting(
  eventId: string,
  options: RespondToMeetingOptions
): Promise<Result<void>> {
  const authResult = requireGraphAuth();
  if (!authResult.ok) return authResult;
  const graphToken = authResult.value;

  const sendResponse = options.sendResponse !== false;
  const body: Record<string, unknown> = { sendResponse };

  // Graph rejects a non-null comment when sendResponse is false, so only
  // include a comment when one was actually provided and we are responding.
  if (options.comment && sendResponse) {
    body.comment = options.comment;
  }

  // Graph accepts a proposed new time on decline and tentativelyAccept, but not accept.
  if (options.proposedNewTime && options.response !== 'accept') {
    body.proposedNewTime = {
      start: { dateTime: options.proposedNewTime.start, timeZone: 'UTC' },
      end: { dateTime: options.proposedNewTime.end, timeZone: 'UTC' },
    };
  }

  const response = await httpRequest<void>(
    GRAPH_CALENDAR_API.respondToEvent(eventId, options.response),
    {
      method: 'POST',
      headers: getGraphHeaders(graphToken),
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) return response;
  return ok(undefined);
}

// ─────────────────────────────────────────────────────────────────────────────
// Free/busy schedule
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Gets the free/busy availability schedule for one or more users.
 *
 * Useful for checking whether attendees are available before scheduling a meeting.
 *
 * @param options - List of email addresses and time window to check
 */
export async function getSchedule(
  options: GetScheduleOptions
): Promise<Result<GetScheduleResult>> {
  const authResult = requireGraphAuth();
  if (!authResult.ok) return authResult;
  const graphToken = authResult.value;

  const body = {
    schedules: options.schedules,
    startTime: { dateTime: options.startTime, timeZone: 'UTC' },
    endTime: { dateTime: options.endTime, timeZone: 'UTC' },
    availabilityViewInterval: options.availabilityViewInterval ?? 30,
  };

  const response = await httpRequest<Record<string, unknown>>(
    GRAPH_CALENDAR_API.getSchedule(),
    {
      method: 'POST',
      headers: { ...getGraphHeaders(graphToken), 'Prefer': GRAPH_UTC_PREFER },
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) return response;

  const data = response.value.data as Record<string, unknown>;
  const rawSchedules = (data.value ?? []) as Array<Record<string, unknown>>;

  const schedules: UserSchedule[] = rawSchedules.map((raw) => {
    const rawItems = (raw.scheduleItems ?? []) as Array<Record<string, unknown>>;
    return {
      scheduleId: raw.scheduleId as string,
      availabilityView: (raw.availabilityView as string) ?? '',
      scheduleItems: rawItems.map((item) => ({
        start: toUtcIso(item.start as { dateTime?: string; timeZone?: string }),
        end: toUtcIso(item.end as { dateTime?: string; timeZone?: string }),
        status: (item.status as ScheduleItem['status']) ?? 'unknown',
        subject: item.subject as string | undefined,
        isPrivate: item.isPrivate as boolean | undefined,
      })),
    };
  });

  return ok({ schedules });
}
