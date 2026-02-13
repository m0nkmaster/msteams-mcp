/**
 * Search-related tool handlers.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolContext, ToolResult } from './index.js';
import { handleApiResult } from './index.js';
import { searchMessages, searchEmails, searchChannels } from '../api/substrate-api.js';
import { getThreadMessages, getConsumptionHorizon, markAsRead } from '../api/chatsvc-api.js';
import {
  DEFAULT_PAGE_SIZE,
  MAX_PAGE_SIZE,
  DEFAULT_THREAD_LIMIT,
  MAX_THREAD_LIMIT,
  DEFAULT_CHANNEL_LIMIT,
  MAX_CHANNEL_LIMIT,
} from '../constants.js';

// ─────────────────────────────────────────────────────────────────────────────
// Schemas
// ─────────────────────────────────────────────────────────────────────────────

export const SearchInputSchema = z.object({
  query: z.string().min(1, 'Query cannot be empty'),
  maxResults: z.number().optional().default(DEFAULT_PAGE_SIZE),
  from: z.number().min(0).optional().default(0),
  size: z.number().min(1).max(MAX_PAGE_SIZE).optional().default(DEFAULT_PAGE_SIZE),
});

export const GetThreadInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  limit: z.number().min(1).max(MAX_THREAD_LIMIT).optional().default(DEFAULT_THREAD_LIMIT),
  markRead: z.boolean().optional().default(false),
  order: z.enum(['asc', 'desc']).optional().default('desc'),
});

export const FindChannelInputSchema = z.object({
  query: z.string().min(1, 'Query cannot be empty'),
  limit: z.number().min(1).max(MAX_CHANNEL_LIMIT).optional().default(DEFAULT_CHANNEL_LIMIT),
});

export const SearchEmailInputSchema = z.object({
  query: z.string().min(1, 'Query cannot be empty'),
  maxResults: z.number().optional().default(DEFAULT_PAGE_SIZE),
  from: z.number().min(0).optional().default(0),
  size: z.number().min(1).max(MAX_PAGE_SIZE).optional().default(DEFAULT_PAGE_SIZE),
});

// ─────────────────────────────────────────────────────────────────────────────
// Tool Definitions
// ─────────────────────────────────────────────────────────────────────────────

const searchToolDefinition: Tool = {
  name: 'teams_search',
  description: 'Search for messages in Microsoft Teams. Returns matching messages with sender, timestamp, content, conversationId (for replies), and pagination info. Supports operators: from:email, to:name, sent:YYYY-MM-DD, sent:today, is:Messages, is:Meetings, is:Channels, is:Chats, hasattachment:true, "Name" for @mentions. The in:channel operator only works reliably when combined with content (e.g., "meeting in:IT Support"). Combine with NOT to exclude. Results sorted by recency.',
  inputSchema: {
    type: 'object',
    properties: {
      query: {
        type: 'string',
        description: 'Search query with optional operators. WORKING: from:email (or name), to:name (spaces not dots - "to:rob macdonald"), sent:YYYY-MM-DD, sent:>=YYYY-MM-DD, sent:today, is:Messages, is:Meetings, is:Channels, is:Chats (case-sensitive, plural required), hasattachment:true, "Display Name" for @mentions. CHANNEL FILTER: in:channel only works reliably WITH content terms (e.g., "budget in:IT Support"). NEVER quote channel names. NOT WORKING: mentions:, sent:lastweek, @me, from:me, is:meeting (must be is:Meetings). Use teams_get_me first to get email/displayName.',
      },
      maxResults: {
        type: 'number',
        description: 'Maximum number of results to return (default: 25)',
      },
      from: {
        type: 'number',
        description: 'Starting offset for pagination (0-based, default: 0). Use this to get subsequent pages of results.',
      },
      size: {
        type: 'number',
        description: 'Page size (default: 25). Number of results per page.',
      },
    },
    required: ['query'],
  },
};

const getThreadToolDefinition: Tool = {
  name: 'teams_get_thread',
  description: 'Get messages from a Teams conversation/thread. Default: newest-first (latest messages at top). For channels: messages include isThreadReply (true for replies) and threadRootId (ID of the post being replied to). Messages without threadRootId are top-level posts. Use threadRootId to group related messages. Each message includes a "when" field with the day of week (e.g., "Friday, January 30, 2026, 10:45 AM UTC"). Returns unread count and can optionally mark as read.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID to get messages from (e.g., "19:abc@thread.tacv2" from search results)',
      },
      limit: {
        type: 'number',
        description: 'Maximum number of messages to return (default: 50, max: 200)',
      },
      markRead: {
        type: 'boolean',
        description: 'If true, marks the conversation as read up to the latest message after fetching (default: false)',
      },
      order: {
        type: 'string',
        enum: ['asc', 'desc'],
        description: 'Sort order: "desc" (newest-first, default) or "asc" (oldest-first for chronological reading)',
      },
    },
    required: ['conversationId'],
  },
};

const findChannelToolDefinition: Tool = {
  name: 'teams_find_channel',
  description: 'Find Teams channels by name. Searches both (1) channels in teams you\'re a member of (reliable) and (2) channels across the organisation (discovery). Results indicate whether you\'re already a member via the isMember field. Channel IDs can be used with teams_get_thread to read messages.',
  inputSchema: {
    type: 'object',
    properties: {
      query: {
        type: 'string',
        description: 'Channel name to search for (partial match)',
      },
      limit: {
        type: 'number',
        description: 'Maximum number of results (default: 10, max: 50)',
      },
    },
    required: ['query'],
  },
};

const searchEmailToolDefinition: Tool = {
  name: 'teams_search_email',
  description: 'Search emails in the user\'s mailbox. Returns matching emails with subject, sender, recipients, timestamp, preview, read status, and pagination info. Supports the same search operators as Outlook: from:email, to:name, subject:"text", hasattachment:true, sent:YYYY-MM-DD, sent:>=YYYY-MM-DD, sent:today, is:read, is:unread. Results sorted by recency. Uses the same authentication as Teams — no additional login required.',
  inputSchema: {
    type: 'object',
    properties: {
      query: {
        type: 'string',
        description: 'Search query with optional operators. Supports: from:email, to:name, subject:"text", hasattachment:true, sent:YYYY-MM-DD, sent:>=YYYY-MM-DD, sent:today, is:read, is:unread. Plain text searches across subject and body.',
      },
      maxResults: {
        type: 'number',
        description: 'Maximum number of results to return (default: 25)',
      },
      from: {
        type: 'number',
        description: 'Starting offset for pagination (0-based, default: 0). Use this to get subsequent pages of results.',
      },
      size: {
        type: 'number',
        description: 'Page size (default: 25). Number of results per page.',
      },
    },
    required: ['query'],
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// Handlers
// ─────────────────────────────────────────────────────────────────────────────

async function handleSearch(
  input: z.infer<typeof SearchInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await searchMessages(input.query, {
    maxResults: input.maxResults,
    from: input.from,
    size: input.size,
  });

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      query: input.query,
      resultCount: result.value.results.length,
      pagination: {
        from: result.value.pagination.from,
        size: result.value.pagination.size,
        returned: result.value.pagination.returned,
        total: result.value.pagination.total,
        hasMore: result.value.pagination.hasMore,
        nextFrom: result.value.pagination.hasMore
          ? result.value.pagination.from + result.value.pagination.returned
          : undefined,
      },
      results: result.value.results,
    },
  };
}

async function handleGetThread(
  input: z.infer<typeof GetThreadInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getThreadMessages(input.conversationId, { 
    limit: input.limit,
    order: input.order,
  });

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  // Get unread status
  let unreadCount: number | undefined;
  let lastReadMessageId: string | undefined;
  
  const horizonResult = await getConsumptionHorizon(input.conversationId);
  if (horizonResult.ok) {
    lastReadMessageId = horizonResult.value.lastReadMessageId;
    
    // Count messages newer than the last-read message
    if (lastReadMessageId) {
      // Find the last-read message to get its timestamp
      const lastReadMsg = result.value.messages.find(m => m.id === lastReadMessageId);
      
      if (lastReadMsg) {
        const lastReadTime = new Date(lastReadMsg.timestamp).getTime();
        unreadCount = result.value.messages.filter(
          m => !m.isFromMe && new Date(m.timestamp).getTime() > lastReadTime
        ).length;
      } else {
        // Last-read message not in our window - it's older, so all messages are unread
        unreadCount = result.value.messages.filter(m => !m.isFromMe).length;
      }
    } else {
      // No consumption horizon means all messages are unread (new conversation)
      unreadCount = result.value.messages.filter(m => !m.isFromMe).length;
    }
  }

  // Mark as read if requested
  let markedAsRead = false;
  if (input.markRead && result.value.messages.length > 0) {
    // Find the latest message - depends on sort order
    // desc (default): newest is first [0], asc: newest is last [-1]
    const latestMessage = input.order === 'asc' 
      ? result.value.messages[result.value.messages.length - 1]
      : result.value.messages[0];
    if (latestMessage) {
      const markResult = await markAsRead(input.conversationId, latestMessage.id);
      markedAsRead = markResult.ok;
    }
  }

  return {
    success: true,
    data: {
      conversationId: result.value.conversationId,
      messageCount: result.value.messages.length,
      unreadCount,
      lastReadMessageId,
      markedAsRead: input.markRead ? markedAsRead : undefined,
      messages: result.value.messages,
    },
  };
}

async function handleSearchEmail(
  input: z.infer<typeof SearchEmailInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await searchEmails(input.query, {
    maxResults: input.maxResults,
    from: input.from,
    size: input.size,
  });

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      query: input.query,
      resultCount: result.value.results.length,
      pagination: {
        from: result.value.pagination.from,
        size: result.value.pagination.size,
        returned: result.value.pagination.returned,
        total: result.value.pagination.total,
        hasMore: result.value.pagination.hasMore,
        nextFrom: result.value.pagination.hasMore
          ? result.value.pagination.from + result.value.pagination.returned
          : undefined,
      },
      ...(result.value.filteredCount ? {
        calendarResponsesFiltered: result.value.filteredCount,
        note: 'Calendar responses (Accepted/Declined/Tentative) were excluded. Each result includes an emailType field.',
      } : {}),
      results: result.value.results,
    },
  };
}

async function handleFindChannel(
  input: z.infer<typeof FindChannelInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await searchChannels(input.query, input.limit);

  return handleApiResult(result, (value) => ({
    query: input.query,
    count: value.returned,
    channels: value.results,
  }));
}

// ─────────────────────────────────────────────────────────────────────────────
// Exports
// ─────────────────────────────────────────────────────────────────────────────

export const searchTool: RegisteredTool<typeof SearchInputSchema> = {
  definition: searchToolDefinition,
  schema: SearchInputSchema,
  handler: handleSearch,
};

export const getThreadTool: RegisteredTool<typeof GetThreadInputSchema> = {
  definition: getThreadToolDefinition,
  schema: GetThreadInputSchema,
  handler: handleGetThread,
};

export const findChannelTool: RegisteredTool<typeof FindChannelInputSchema> = {
  definition: findChannelToolDefinition,
  schema: FindChannelInputSchema,
  handler: handleFindChannel,
};

export const searchEmailTool: RegisteredTool<typeof SearchEmailInputSchema> = {
  definition: searchEmailToolDefinition,
  schema: SearchEmailInputSchema,
  handler: handleSearchEmail,
};

/** All search-related tools. */
export const searchTools = [searchTool, getThreadTool, findChannelTool, searchEmailTool];
