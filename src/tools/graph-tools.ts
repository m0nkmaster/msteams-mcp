/**
 * Graph API spike tool handlers.
 * 
 * Experimental tools to test Microsoft Graph API access using tokens
 * extracted from the Teams browser session.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolContext, ToolResult } from './index.js';
import { graphSendMessage } from '../api/graph-api.js';
import { getGraphTokenStatus } from '../auth/token-extractor.js';

// ─────────────────────────────────────────────────────────────────────────────
// Schemas
// ─────────────────────────────────────────────────────────────────────────────

export const GraphSendMessageInputSchema = z.object({
  chatId: z.string().min(1, 'Chat ID cannot be empty'),
  content: z.string().min(1, 'Message content cannot be empty'),
  contentType: z.enum(['text', 'html']).optional().default('text'),
});

export const GraphTokenStatusInputSchema = z.object({});

// ─────────────────────────────────────────────────────────────────────────────
// Tool Definitions
// ─────────────────────────────────────────────────────────────────────────────

const graphSendMessageToolDefinition: Tool = {
  name: 'teams_graph_send_message',
  description: '[SPIKE] Send a message via Microsoft Graph API instead of chatsvc. This is an experimental tool to test Graph API access. Use the chatId from teams_get_thread or teams_search (the conversationId). Content can be plain text or HTML.',
  inputSchema: {
    type: 'object',
    properties: {
      chatId: {
        type: 'string',
        description: 'The Teams chat/conversation ID (e.g., "19:xxx@thread.v2"). Use the conversationId from other tools.',
      },
      content: {
        type: 'string',
        description: 'The message content. Plain text by default, or HTML if contentType is "html".',
      },
      contentType: {
        type: 'string',
        enum: ['text', 'html'],
        description: 'Content type: "text" (default) or "html".',
      },
    },
    required: ['chatId', 'content'],
  },
};

const graphTokenStatusToolDefinition: Tool = {
  name: 'teams_graph_token_status',
  description: '[SPIKE] Check the status of the Microsoft Graph API token. Shows whether a Graph token is available and when it expires.',
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// Handlers
// ─────────────────────────────────────────────────────────────────────────────

async function handleGraphSendMessage(
  input: z.infer<typeof GraphSendMessageInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await graphSendMessage(
    input.chatId,
    input.content,
    input.contentType
  );

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      id: result.value.id,
      createdDateTime: result.value.createdDateTime,
      webUrl: result.value.webUrl,
      note: '[SPIKE] Message sent via Microsoft Graph API. This is experimental.',
      _raw: result.value._raw,
    },
  };
}

async function handleGraphTokenStatus(
  _input: Record<string, never>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const status = getGraphTokenStatus();

  return {
    success: true,
    data: {
      graphToken: status,
      note: '[SPIKE] Graph token is acquired during token refresh. If no token is available, try teams_login first, then check again.',
    },
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Exports
// ─────────────────────────────────────────────────────────────────────────────

export const graphSendMessageTool: RegisteredTool<typeof GraphSendMessageInputSchema> = {
  definition: graphSendMessageToolDefinition,
  schema: GraphSendMessageInputSchema,
  handler: handleGraphSendMessage,
};

export const graphTokenStatusTool: RegisteredTool<z.ZodObject<Record<string, never>>> = {
  definition: graphTokenStatusToolDefinition,
  schema: z.object({}),
  handler: handleGraphTokenStatus,
};

/** All Graph API spike tools. */
export const graphTools = [
  graphSendMessageTool,
  graphTokenStatusTool,
];
