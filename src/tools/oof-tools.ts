/**
 * Out-of-office (automatic replies) tool handlers.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolResult } from './index.js';
import { handleApiResult } from './index.js';
import { getAutomaticReplies, setAutomaticReplies } from '../api/oof-api.js';

// ─────────────────────────────────────────────────────────────────────────────
// Schemas
// ─────────────────────────────────────────────────────────────────────────────

export const GetOutOfOfficeInputSchema = z.object({});

export const SetOutOfOfficeInputSchema = z.object({
  enabled: z.boolean(),
  internalMessage: z.string().optional(),
  externalMessage: z.string().optional(),
  externalAudience: z.enum(['none', 'contactsOnly', 'all']).optional(),
  until: z.string().datetime({ offset: true }).optional(),
});

// ─────────────────────────────────────────────────────────────────────────────
// Tool Definitions
// ─────────────────────────────────────────────────────────────────────────────

const getOutOfOfficeToolDefinition: Tool = {
  name: 'teams_get_out_of_office',
  description: 'Get the current out-of-office (Outlook automatic replies) status. Returns status ("disabled" | "alwaysEnabled" | "scheduled"), the internal/external reply messages, externalAudience, and the scheduled window when set. This is the same OOF that Teams shows as your out-of-office note.',
  inputSchema: {
    type: 'object',
    properties: {},
    required: [],
  },
};

const setOutOfOfficeToolDefinition: Tool = {
  name: 'teams_set_out_of_office',
  description: 'Turn out-of-office (Outlook automatic replies) ON or OFF. This is the canonical OOF that Teams reflects automatically. Set enabled=true with a message to turn on; enabled=false to turn off. Provide until (ISO 8601) to auto-expire at a time (scheduled mode), otherwise it stays on until turned off.',
  inputSchema: {
    type: 'object',
    properties: {
      enabled: {
        type: 'boolean',
        description: 'true to turn automatic replies ON, false to turn them OFF.',
      },
      internalMessage: {
        type: 'string',
        description: 'The reply text sent to internal senders (also used for external senders unless externalMessage is given). Optional; when omitted the auto-reply is left empty. Note: use --internalMessage on the CLI (--message is reserved for message IDs).',
      },
      externalMessage: {
        type: 'string',
        description: 'Optional separate reply for external senders. Defaults to the internal message.',
      },
      externalAudience: {
        type: 'string',
        enum: ['none', 'contactsOnly', 'all'],
        description: 'Who receives the external reply. Defaults to "all".',
      },
      until: {
        type: 'string',
        description: 'Optional ISO 8601 end time (e.g. "2026-09-22T07:00:00Z"). When set, uses scheduled mode (starts now, ends then).',
      },
    },
    required: ['enabled'],
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// Handlers
// ─────────────────────────────────────────────────────────────────────────────

async function handleGetOutOfOffice(): Promise<ToolResult> {
  const result = await getAutomaticReplies();

  return handleApiResult(result, (value) => ({
    status: value.status,
    enabled: value.status !== 'disabled',
    externalAudience: value.externalAudience,
    internalReplyMessage: value.internalReplyMessage,
    externalReplyMessage: value.externalReplyMessage,
    scheduledStartDateTime: value.scheduledStartDateTime,
    scheduledEndDateTime: value.scheduledEndDateTime,
  }));
}

async function handleSetOutOfOffice(
  input: z.infer<typeof SetOutOfOfficeInputSchema>
): Promise<ToolResult> {
  const result = await setAutomaticReplies(input);

  return handleApiResult(result, (value) => ({
    status: value.status,
    enabled: value.status !== 'disabled',
    externalAudience: value.externalAudience,
    scheduledEndDateTime: value.scheduledEndDateTime,
  }));
}

// ─────────────────────────────────────────────────────────────────────────────
// Exports
// ─────────────────────────────────────────────────────────────────────────────

export const getOutOfOfficeTool: RegisteredTool<typeof GetOutOfOfficeInputSchema> = {
  definition: getOutOfOfficeToolDefinition,
  schema: GetOutOfOfficeInputSchema,
  handler: handleGetOutOfOffice,
};

export const setOutOfOfficeTool: RegisteredTool<typeof SetOutOfOfficeInputSchema> = {
  definition: setOutOfOfficeToolDefinition,
  schema: SetOutOfOfficeInputSchema,
  handler: handleSetOutOfOffice,
};

/** All out-of-office tools. */
export const outOfOfficeTools = [getOutOfOfficeTool, setOutOfOfficeTool];
