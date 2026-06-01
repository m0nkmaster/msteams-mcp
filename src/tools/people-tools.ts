/**
 * People-related tool handlers.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolContext, ToolResult } from './index.js';
import { handleApiResult } from './index.js';
import { searchPeople, getFrequentContacts } from '../api/substrate-api.js';
import { getPresence } from '../api/presence-api.js';
import { getUserProfile } from '../auth/token-extractor.js';
import { ErrorCode, createError } from '../types/errors.js';
import {
  DEFAULT_PEOPLE_LIMIT,
  MAX_PEOPLE_LIMIT,
  DEFAULT_CONTACTS_LIMIT,
  MAX_CONTACTS_LIMIT,
} from '../constants.js';

// ─────────────────────────────────────────────────────────────────────────────
// Schemas
// ─────────────────────────────────────────────────────────────────────────────

export const SearchPeopleInputSchema = z.object({
  query: z.string().min(1, 'Query cannot be empty'),
  limit: z.number().min(1).max(MAX_PEOPLE_LIMIT).optional().default(DEFAULT_PEOPLE_LIMIT),
});

export const FrequentContactsInputSchema = z.object({
  limit: z.number().min(1).max(MAX_CONTACTS_LIMIT).optional().default(DEFAULT_CONTACTS_LIMIT),
});

export const GetPresenceInputSchema = z.object({
  userIds: z.array(z.string().min(1)).min(1),
});

// ─────────────────────────────────────────────────────────────────────────────
// Tool Definitions
// ─────────────────────────────────────────────────────────────────────────────

const getMeToolDefinition: Tool = {
  name: 'teams_get_me',
  description: 'Get the current user\'s profile information including email, display name, and Teams ID. Useful for finding @mentions or identifying the current user.',
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

const searchPeopleToolDefinition: Tool = {
  name: 'teams_search_people',
  description: 'Search for people in Microsoft Teams by name or email. Returns matching users with display name, email, job title, and department. Useful for finding someone to message.',
  inputSchema: {
    type: 'object',
    properties: {
      query: {
        type: 'string',
        description: 'Search term - can be a name, email address, or partial match',
      },
      limit: {
        type: 'number',
        description: 'Maximum number of results to return (default: 10)',
      },
    },
    required: ['query'],
  },
};

const frequentContactsToolDefinition: Tool = {
  name: 'teams_get_frequent_contacts',
  description: 'Get the user\'s frequently contacted people, ranked by interaction frequency. Useful for resolving ambiguous names (e.g., "Rob" → which Rob?) by checking who the user commonly works with. Returns display name, email, job title, and department.',
  inputSchema: {
    type: 'object',
    properties: {
      limit: {
        type: 'number',
        description: 'Maximum number of contacts to return (default: 50)',
      },
    },
  },
};

const getPresenceToolDefinition: Tool = {
  name: 'teams_get_presence',
  description: 'Get the real time presence of one or more people: availability (Available, Busy, Away, Offline, DoNotDisturb), activity (e.g. InACall, InAMeeting, Presenting), device, last active time, whether they are out of office, and their out of office auto reply message if set. Pass user identifiers (MRI like "8:orgid:<guid>" or a raw object id from teams_search_people / teams_get_frequent_contacts).',
  inputSchema: {
    type: 'object',
    properties: {
      userIds: {
        type: 'array',
        items: { type: 'string' },
        description: 'User identifiers to look up: MRI ("8:orgid:<guid>") or raw object id. Get these from teams_search_people or teams_get_frequent_contacts.',
      },
    },
    required: ['userIds'],
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// Handlers
// ─────────────────────────────────────────────────────────────────────────────

async function handleGetMe(
  _input: Record<string, never>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const profile = getUserProfile();

  if (!profile) {
    return {
      success: false,
      error: createError(
        ErrorCode.AUTH_REQUIRED,
        'No valid session. Please use teams_login first.'
      ),
    };
  }

  return {
    success: true,
    data: { profile },
  };
}

async function handleSearchPeople(
  input: z.infer<typeof SearchPeopleInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await searchPeople(input.query, input.limit);

  return handleApiResult(result, (value) => ({
    query: input.query,
    returned: value.returned,
    results: value.results,
  }));
}

async function handleGetFrequentContacts(
  input: z.infer<typeof FrequentContactsInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getFrequentContacts(input.limit);

  return handleApiResult(result, (value) => ({
    returned: value.returned,
    contacts: value.results,
  }));
}

async function handleGetPresence(
  input: z.infer<typeof GetPresenceInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getPresence(input.userIds);

  return handleApiResult(result, (value) => ({ presences: value }));
}

// ─────────────────────────────────────────────────────────────────────────────
// Exports
// ─────────────────────────────────────────────────────────────────────────────

export const getMeTool: RegisteredTool<z.ZodObject<Record<string, never>>> = {
  definition: getMeToolDefinition,
  schema: z.object({}),
  handler: handleGetMe,
};

export const searchPeopleTool: RegisteredTool<typeof SearchPeopleInputSchema> = {
  definition: searchPeopleToolDefinition,
  schema: SearchPeopleInputSchema,
  handler: handleSearchPeople,
};

export const frequentContactsTool: RegisteredTool<typeof FrequentContactsInputSchema> = {
  definition: frequentContactsToolDefinition,
  schema: FrequentContactsInputSchema,
  handler: handleGetFrequentContacts,
};

export const getPresenceTool: RegisteredTool<typeof GetPresenceInputSchema> = {
  definition: getPresenceToolDefinition,
  schema: GetPresenceInputSchema,
  handler: handleGetPresence,
};

/** All people-related tools. */
export const peopleTools = [getMeTool, searchPeopleTool, frequentContactsTool, getPresenceTool];
