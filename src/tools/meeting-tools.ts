/**
 * Meeting/calendar-related tool handlers.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolContext, ToolResult } from './index.js';
import { handleApiResult } from './index.js';
import { getCalendarView } from '../api/calendar-api.js';
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
// Exports
// ─────────────────────────────────────────────────────────────────────────────

export const getMeetingsTool: RegisteredTool<typeof GetMeetingsInputSchema> = {
  definition: getMeetingsToolDefinition,
  schema: GetMeetingsInputSchema,
  handler: handleGetMeetings,
};

/** All meeting-related tools. */
export const meetingTools = [getMeetingsTool];
