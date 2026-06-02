/**
 * Meeting/calendar-related tool handlers.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolContext, ToolResult } from './index.js';
import { handleApiResult } from './index.js';
import {
  getCalendarView,
  createMeeting,
  getMeeting,
  updateMeeting,
  cancelMeeting,
  respondToMeeting,
  getSchedule,
} from '../api/calendar-api.js';
import { getTranscriptContent } from '../api/transcript-api.js';
import {
  DEFAULT_MEETING_LIMIT,
  MAX_MEETING_LIMIT,
} from '../constants.js';

// ─────────────────────────────────────────────────────────────────────────────
// Schemas
// ─────────────────────────────────────────────────────────────────────────────

export const GetMeetingsInputSchema = z.object({
  startDate: z.string().optional(),
  endDate: z.string().optional(),
  limit: z.number().min(1).max(MAX_MEETING_LIMIT).optional().default(DEFAULT_MEETING_LIMIT),
});

export const GetTranscriptInputSchema = z.object({
  threadId: z.string().min(1),
  meetingDate: z.string().optional(),
});

export const CreateMeetingInputSchema = z.object({
  subject: z.string().min(1),
  startTime: z.string().min(1),
  endTime: z.string().min(1),
  attendees: z.array(z.object({
    email: z.string().email(),
    name: z.string().optional(),
  })).optional(),
  body: z.string().optional(),
  isOnlineMeeting: z.boolean().optional().default(true),
  location: z.string().optional(),
});

export const GetMeetingInputSchema = z.object({
  eventId: z.string().min(1),
});

export const UpdateMeetingInputSchema = z.object({
  eventId: z.string().min(1),
  subject: z.string().min(1).optional(),
  startTime: z.string().min(1).optional(),
  endTime: z.string().min(1).optional(),
  attendees: z.array(z.object({
    email: z.string().email(),
    name: z.string().optional(),
  })).optional(),
  body: z.string().optional(),
  location: z.string().optional(),
});

export const CancelMeetingInputSchema = z.object({
  eventId: z.string().min(1),
});

export const RespondToMeetingInputSchema = z.object({
  eventId: z.string().min(1),
  response: z.enum(['accept', 'tentativelyAccept', 'decline']),
  comment: z.string().optional(),
  sendResponse: z.boolean().optional().default(true),
  proposedNewTime: z.object({
    start: z.string().min(1),
    end: z.string().min(1),
  }).optional(),
});

export const GetScheduleInputSchema = z.object({
  schedules: z.array(z.string().email()).min(1),
  startTime: z.string().min(1),
  endTime: z.string().min(1),
  availabilityViewInterval: z.number().min(5).max(1440).optional().default(30),
});

// ─────────────────────────────────────────────────────────────────────────────
// Tool Definitions
// ─────────────────────────────────────────────────────────────────────────────

const getMeetingsToolDefinition: Tool = {
  name: 'teams_get_meetings',
  description: 'Get meetings from your Teams calendar. Returns meetings with: subject, startTime, endTime, organizer (name/email), location, joinUrl (Teams link), threadId (use with teams_get_thread to read meeting chat), myResponse (None/Accepted/Tentative/Declined), showAs (Free/Busy), isOrganizer. Defaults to next 7 days from now. For past meetings (e.g., "summarise my last meeting"), set startDate to a past date. To find meetings with a person, get results and filter by organizer.email.',
  inputSchema: {
    type: 'object',
    properties: {
      startDate: {
        type: 'string',
        description: 'Start of date range (ISO 8601, e.g., "2026-02-01T00:00:00.000Z"). Defaults to now.',
      },
      endDate: {
        type: 'string',
        description: 'End of date range (ISO 8601). Defaults to 7 days from now.',
      },
      limit: {
        type: 'number',
        description: 'Maximum number of meetings to return (default: 50, max: 200)',
      },
    },
    required: [],
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// Handlers
// ─────────────────────────────────────────────────────────────────────────────

async function handleGetMeetings(
  input: z.infer<typeof GetMeetingsInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getCalendarView({
    startDate: input.startDate,
    endDate: input.endDate,
    limit: input.limit,
  });

  return handleApiResult(result, (value) => ({
    count: value.count,
    meetings: value.meetings,
  }));
}

// ─────────────────────────────────────────────────────────────────────────────
// Create Meeting Tool
// ─────────────────────────────────────────────────────────────────────────────

const createMeetingToolDefinition: Tool = {
  name: 'teams_create_meeting',
  description: 'Create a new Teams calendar meeting and send invites to attendees. Returns the meeting ID, subject, start/end times, and the Teams join URL. Set isOnlineMeeting to true (default) to generate a Teams meeting link. Attendees receive calendar invites automatically.',
  inputSchema: {
    type: 'object',
    properties: {
      subject: {
        type: 'string',
        description: 'Meeting title/subject.',
      },
      startTime: {
        type: 'string',
        description: 'Start time in ISO 8601 format (e.g., "2026-06-05T10:00:00Z"). Use UTC.',
      },
      endTime: {
        type: 'string',
        description: 'End time in ISO 8601 format (e.g., "2026-06-05T10:30:00Z"). Use UTC.',
      },
      attendees: {
        type: 'array',
        description: 'List of attendees to invite.',
        items: {
          type: 'object',
          properties: {
            email: { type: 'string', description: 'Attendee email address.' },
            name: { type: 'string', description: 'Attendee display name (optional).' },
          },
          required: ['email'],
        },
      },
      body: {
        type: 'string',
        description: 'Meeting description or agenda (plain text).',
      },
      isOnlineMeeting: {
        type: 'boolean',
        description: 'Whether to generate a Teams meeting link (default: true).',
      },
      location: {
        type: 'string',
        description: 'Meeting location (room name or free text, optional).',
      },
    },
    required: ['subject', 'startTime', 'endTime'],
  },
};

async function handleCreateMeeting(
  input: z.infer<typeof CreateMeetingInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await createMeeting({
    subject: input.subject,
    startTime: input.startTime,
    endTime: input.endTime,
    attendees: input.attendees,
    body: input.body,
    isOnlineMeeting: input.isOnlineMeeting,
    location: input.location,
  });

  return handleApiResult(result, (value) => ({
    id: value.id,
    subject: value.subject,
    startTime: value.startTime,
    endTime: value.endTime,
    joinUrl: value.joinUrl,
  }));
}

// ─────────────────────────────────────────────────────────────────────────────
// Get single meeting tool
// ─────────────────────────────────────────────────────────────────────────────

const getMeetingToolDefinition: Tool = {
  name: 'teams_get_meeting',
  description: 'Get full details of a single Teams calendar event by its ID. Returns all fields including the full attendee list, meeting body/description, location, and Teams join URL. Use the eventId from teams_get_meetings results.',
  inputSchema: {
    type: 'object',
    properties: {
      eventId: {
        type: 'string',
        description: 'The event ID (the id field from teams_get_meetings results).',
      },
    },
    required: ['eventId'],
  },
};

async function handleGetMeeting(
  input: z.infer<typeof GetMeetingInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getMeeting(input.eventId);
  return handleApiResult(result, (value) => ({ meeting: value }));
}

// ─────────────────────────────────────────────────────────────────────────────
// Update meeting tool
// ─────────────────────────────────────────────────────────────────────────────

const updateMeetingToolDefinition: Tool = {
  name: 'teams_update_meeting',
  description: 'Update (reschedule or edit) an existing Teams calendar event. Only the fields you provide are changed — omit any field to leave it unchanged. Returns the updated meeting. You must be the organiser to update the event for all attendees.',
  inputSchema: {
    type: 'object',
    properties: {
      eventId: {
        type: 'string',
        description: 'The event ID to update (from teams_get_meetings).',
      },
      subject: { type: 'string', description: 'New meeting title.' },
      startTime: { type: 'string', description: 'New start time (ISO 8601 UTC).' },
      endTime: { type: 'string', description: 'New end time (ISO 8601 UTC).' },
      attendees: {
        type: 'array',
        description: 'Replaces the full attendee list.',
        items: {
          type: 'object',
          properties: {
            email: { type: 'string' },
            name: { type: 'string' },
          },
          required: ['email'],
        },
      },
      body: { type: 'string', description: 'New meeting description / agenda (plain text).' },
      location: { type: 'string', description: 'New meeting location.' },
    },
    required: ['eventId'],
  },
};

async function handleUpdateMeeting(
  input: z.infer<typeof UpdateMeetingInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const { eventId, ...options } = input;
  const result = await updateMeeting(eventId, options);
  return handleApiResult(result, (value) => ({ meeting: value }));
}

// ─────────────────────────────────────────────────────────────────────────────
// Cancel meeting tool
// ─────────────────────────────────────────────────────────────────────────────

const cancelMeetingToolDefinition: Tool = {
  name: 'teams_cancel_meeting',
  description: 'Cancel or remove a Teams calendar event. If you are the organiser, this sends a cancellation notice to all attendees. If you are an attendee, this removes it from your calendar only without notifying others.',
  inputSchema: {
    type: 'object',
    properties: {
      eventId: {
        type: 'string',
        description: 'The event ID to cancel (from teams_get_meetings).',
      },
    },
    required: ['eventId'],
  },
};

async function handleCancelMeeting(
  input: z.infer<typeof CancelMeetingInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await cancelMeeting(input.eventId);
  return handleApiResult(result, () => ({ success: true }));
}

// ─────────────────────────────────────────────────────────────────────────────
// Respond to meeting tool
// ─────────────────────────────────────────────────────────────────────────────

const respondToMeetingToolDefinition: Tool = {
  name: 'teams_respond_to_meeting',
  description: 'Accept, tentatively accept, or decline a Teams meeting invite. Optionally include a comment to send to the organiser (the comment is only sent when sendResponse is true). When declining or tentatively accepting, you can propose an alternative time. Set sendResponse to false to update your calendar silently without emailing the organiser.',
  inputSchema: {
    type: 'object',
    properties: {
      eventId: {
        type: 'string',
        description: 'The event ID to respond to (from teams_get_meetings).',
      },
      response: {
        type: 'string',
        enum: ['accept', 'tentativelyAccept', 'decline'],
        description: '"accept" to accept, "tentativelyAccept" for maybe, "decline" to decline.',
      },
      comment: {
        type: 'string',
        description: 'Optional message to send with your response.',
      },
      sendResponse: {
        type: 'boolean',
        description: 'Whether to notify the organiser (default: true). Set to false to update your calendar silently.',
      },
      proposedNewTime: {
        type: 'object',
        description: 'When declining or tentatively accepting, propose an alternative time (only if the event allows it). Ignored when accepting.',
        properties: {
          start: { type: 'string', description: 'Proposed start time (ISO 8601 UTC).' },
          end: { type: 'string', description: 'Proposed end time (ISO 8601 UTC).' },
        },
        required: ['start', 'end'],
      },
    },
    required: ['eventId', 'response'],
  },
};

async function handleRespondToMeeting(
  input: z.infer<typeof RespondToMeetingInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await respondToMeeting(input.eventId, {
    response: input.response,
    comment: input.comment,
    sendResponse: input.sendResponse,
    proposedNewTime: input.proposedNewTime,
  });
  return handleApiResult(result, () => ({ success: true }));
}

// ─────────────────────────────────────────────────────────────────────────────
// Get schedule (free/busy) tool
// ─────────────────────────────────────────────────────────────────────────────

const getScheduleToolDefinition: Tool = {
  name: 'teams_get_schedule',
  description: 'Check the free/busy availability of one or more people for a given time window. Returns each person\'s busy slots and a visual availability string (0=free, 1=tentative, 2=busy, 3=OOF) in configurable intervals. Useful for finding a good meeting time before sending an invite.',
  inputSchema: {
    type: 'object',
    properties: {
      schedules: {
        type: 'array',
        items: { type: 'string' },
        description: 'Email addresses of the people to check (e.g., ["chris@company.com", "atakan@company.com"]).',
      },
      startTime: {
        type: 'string',
        description: 'Start of the window to check (ISO 8601 UTC, e.g., "2026-06-05T09:00:00Z").',
      },
      endTime: {
        type: 'string',
        description: 'End of the window to check (ISO 8601 UTC, e.g., "2026-06-05T18:00:00Z").',
      },
      availabilityViewInterval: {
        type: 'number',
        description: 'Slot size in minutes for the availability string (default: 30, min: 5, max: 1440).',
      },
    },
    required: ['schedules', 'startTime', 'endTime'],
  },
};

async function handleGetSchedule(
  input: z.infer<typeof GetScheduleInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getSchedule({
    schedules: input.schedules,
    startTime: input.startTime,
    endTime: input.endTime,
    availabilityViewInterval: input.availabilityViewInterval,
  });
  return handleApiResult(result, (value) => ({ schedules: value.schedules }));
}

// ─────────────────────────────────────────────────────────────────────────────
// Transcript Tool
// ─────────────────────────────────────────────────────────────────────────────

const getTranscriptToolDefinition: Tool = {
  name: 'teams_get_transcript',
  description: 'Get the transcript of a Teams meeting. Requires the meeting\'s threadId (from teams_get_meetings). Returns formatted transcript text with timestamps and speaker names, ready for reading or summarization. The meeting must have had transcription enabled. Optionally pass meetingDate (ISO string, e.g. the startTime from teams_get_meetings) to narrow the search.',
  inputSchema: {
    type: 'object',
    properties: {
      threadId: {
        type: 'string',
        description: 'The meeting thread ID (from the threadId field of teams_get_meetings results, e.g., "19:meeting_xxx@thread.v2").',
      },
      meetingDate: {
        type: 'string',
        description: 'Optional ISO date/time of the meeting (e.g., the startTime from teams_get_meetings). Helps narrow the search for recurring meetings.',
      },
    },
    required: ['threadId'],
  },
};

async function handleGetTranscript(
  input: z.infer<typeof GetTranscriptInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getTranscriptContent(input.threadId, input.meetingDate);

  return handleApiResult(result, (value) => ({
    meetingTitle: value.meetingTitle,
    speakers: value.speakers,
    entryCount: value.entryCount,
    transcript: value.formattedText,
  }));
}

// ─────────────────────────────────────────────────────────────────────────────
// Exports
// ─────────────────────────────────────────────────────────────────────────────

export const getMeetingsTool: RegisteredTool<typeof GetMeetingsInputSchema> = {
  definition: getMeetingsToolDefinition,
  schema: GetMeetingsInputSchema,
  handler: handleGetMeetings,
};

export const createMeetingTool: RegisteredTool<typeof CreateMeetingInputSchema> = {
  definition: createMeetingToolDefinition,
  schema: CreateMeetingInputSchema,
  handler: handleCreateMeeting,
};

export const getMeetingTool: RegisteredTool<typeof GetMeetingInputSchema> = {
  definition: getMeetingToolDefinition,
  schema: GetMeetingInputSchema,
  handler: handleGetMeeting,
};

export const updateMeetingTool: RegisteredTool<typeof UpdateMeetingInputSchema> = {
  definition: updateMeetingToolDefinition,
  schema: UpdateMeetingInputSchema,
  handler: handleUpdateMeeting,
};

export const cancelMeetingTool: RegisteredTool<typeof CancelMeetingInputSchema> = {
  definition: cancelMeetingToolDefinition,
  schema: CancelMeetingInputSchema,
  handler: handleCancelMeeting,
};

export const respondToMeetingTool: RegisteredTool<typeof RespondToMeetingInputSchema> = {
  definition: respondToMeetingToolDefinition,
  schema: RespondToMeetingInputSchema,
  handler: handleRespondToMeeting,
};

export const getScheduleTool: RegisteredTool<typeof GetScheduleInputSchema> = {
  definition: getScheduleToolDefinition,
  schema: GetScheduleInputSchema,
  handler: handleGetSchedule,
};

export const getTranscriptTool: RegisteredTool<typeof GetTranscriptInputSchema> = {
  definition: getTranscriptToolDefinition,
  schema: GetTranscriptInputSchema,
  handler: handleGetTranscript,
};

/** All meeting-related tools. */
export const meetingTools = [
  getMeetingsTool,
  createMeetingTool,
  getMeetingTool,
  updateMeetingTool,
  cancelMeetingTool,
  respondToMeetingTool,
  getScheduleTool,
  getTranscriptTool,
];
