/**
 * Tag-related tool handlers.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolContext, ToolResult } from './index.js';
import { listTeamTags } from '../api/tags-api.js';
import { handleApiResult } from './index.js';

// ─────────────────────────────────────────────────────────────────────────────
// Schemas
// ─────────────────────────────────────────────────────────────────────────────

export const GetTagsInputSchema = z.object({
  teamId: z.string().min(1, 'Team ID cannot be empty'),
  query: z.string().optional(),
});

// ─────────────────────────────────────────────────────────────────────────────
// Tool Definitions
// ─────────────────────────────────────────────────────────────────────────────

const getTagsDefinition: Tool = {
  name: 'teams_get_tags',
  description: `Get tags for a Teams team. Returns tag IDs, display names, and member counts. Use the teamId from teams_find_channel results. Optionally filter by name with the query parameter.

To @mention a tag in a message, use the tag MRI format: @[TagName](tag:{tagId}). Example: after finding tag "engineering" with id "abc123", use @[engineering](tag:abc123) in teams_send_message content.`,
  inputSchema: {
    type: 'object' as const,
    properties: {
      teamId: {
        type: 'string',
        description: 'The team group ID (GUID). Get this from teams_find_channel results (teamId field).',
      },
      query: {
        type: 'string',
        description: 'Optional: filter tags by name (case-insensitive partial match).',
      },
    },
    required: ['teamId'],
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// Handlers
// ─────────────────────────────────────────────────────────────────────────────

async function handleGetTags(
  input: z.infer<typeof GetTagsInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await listTeamTags(input.teamId);

  return handleApiResult(result, (value) => {
    let tags = value.tags;

    // Filter by query if provided
    if (input.query) {
      const q = input.query.toLowerCase();
      tags = tags.filter(t => t.displayName.toLowerCase().includes(q));
    }

    return {
      success: true,
      teamId: input.teamId,
      count: tags.length,
      tags: tags.map(t => ({
        id: t.id,
        displayName: t.displayName,
        memberCount: t.memberCount,
        tagType: t.tagType,
        mentionSyntax: `@[${t.displayName}](tag:${t.id})`,
      })),
      ...(input.query ? { query: input.query } : {}),
      note: 'Use the mentionSyntax value in teams_send_message content to @mention a tag.',
    };
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Registration
// ─────────────────────────────────────────────────────────────────────────────

const getTagsTool: RegisteredTool<typeof GetTagsInputSchema> = {
  definition: getTagsDefinition,
  schema: GetTagsInputSchema,
  handler: handleGetTags,
};

export const tagTools = [getTagsTool];
