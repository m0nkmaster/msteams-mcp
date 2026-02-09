/**
 * People-related tool handlers.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolContext, ToolResult } from './index.js';
import { handleApiResult } from './index.js';
import { searchPeople, getFrequentContacts } from '../api/substrate-api.js';
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

/** All people-related tools. */
export const peopleTools = [getMeTool, searchPeopleTool, frequentContactsTool];
