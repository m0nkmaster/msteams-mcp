/**
 * Calendar API client for meeting operations.
 * 
 * Handles calls to teams.microsoft.com/api/mt/part/{region}/v2.1/me/calendars endpoints.
 */

import { httpRequest } from '../utils/http.js';
import { CALENDAR_API, getTeamsHeaders } from '../utils/api-config.js';
import { type Result, ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import { requireSkypeSpacesAuth, getRegionConfig } from '../utils/auth-guards.js';
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

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

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
    myResponse,
    showAs,
    isOrganizer: raw.isOrganizer === true,
    eventType: (raw.eventType as string) || 'Single',
  };
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
 * Uses the Teams mt/part API with the same Skype Spaces auth as getCalendarView.
 * Optionally adds a Teams online meeting link and sends invites to attendees.
 *
 * @param options - Event creation options
 * @returns The created meeting's ID, subject, times, and join URL
 */
export async function createMeeting(
  options: CreateMeetingOptions
): Promise<Result<CreatedMeeting>> {
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

  const body: Record<string, unknown> = {
    subject: options.subject,
    start: { dateTime: options.startTime, timeZone: 'UTC' },
    end: { dateTime: options.endTime, timeZone: 'UTC' },
    isOnlineMeeting: options.isOnlineMeeting !== false,
    onlineMeetingProvider: 'teamsForBusiness',
  };

  if (options.attendees && options.attendees.length > 0) {
    body.attendees = options.attendees.map((a) => ({
      emailAddress: { address: a.email, name: a.name ?? a.email },
      type: 'required',
    }));
  }

  if (options.body) {
    body.body = { contentType: 'text', content: options.body };
  }

  if (options.location) {
    body.location = { displayName: options.location };
  }

  const createUrl = CALENDAR_API.createEvent(
    regionConfig.regionPartition,
    regionConfig.hasPartition,
    regionConfig.teamsBaseUrl
  );

  const response = await httpRequest<Record<string, unknown>>(
    createUrl,
    {
      method: 'POST',
      headers: {
        ...getTeamsHeaders(regionConfig.teamsBaseUrl),
        'Authentication': `skypetoken=${skypeToken}`,
        'Authorization': `Bearer ${spacesToken}`,
      },
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    return response;
  }

  const data = response.value.data as Record<string, unknown>;

  // Extract the Teams join URL from the online meeting details
  const onlineMeeting = data.onlineMeeting as Record<string, unknown> | undefined;
  const joinUrl = (onlineMeeting?.joinUrl as string | undefined)
    ?? (data.skypeTeamsMeetingUrl as string | undefined);

  return ok({
    id: data.id as string,
    subject: (data.subject as string) ?? options.subject,
    startTime: (data.start as Record<string, string>)?.dateTime ?? options.startTime,
    endTime: (data.end as Record<string, string>)?.dateTime ?? options.endTime,
    joinUrl,
  });
}
