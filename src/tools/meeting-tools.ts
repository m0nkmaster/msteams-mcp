/**
 * Meeting/calendar-related tool handlers.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolContext, ToolResult } from './index.js';
import { handleApiResult } from './index.js';
import { getCalendarView } from '../api/calendar-api.js';
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
  threadId: z.string(),
  meetingDate: z.string().optional(),
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

export const getTranscriptTool: RegisteredTool<typeof GetTranscriptInputSchema> = {
  definition: getTranscriptToolDefinition,
  schema: GetTranscriptInputSchema,
  handler: handleGetTranscript,
};

/** All meeting-related tools. */
export const meetingTools = [getMeetingsTool, getTranscriptTool];
