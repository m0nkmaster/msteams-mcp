/**
 * Messaging-related tool handlers.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolContext, ToolResult } from './index.js';
import {
  sendMessage,
  saveMessage,
  unsaveMessage,
  getOneOnOneChatId,
  createGroupChat,
  editMessage,
  deleteMessage,
  getUnreadStatus,
  markAsRead,
  getActivityFeed,
  addReaction,
  removeReaction,
  getSavedMessages,
  getFollowedThreads,
} from '../api/chatsvc-api.js';
import { getFavorites, addFavorite, removeFavorite, getCustomEmojis } from '../api/csa-api.js';
import { SELF_CHAT_ID, MAX_UNREAD_AGGREGATE_CHECK, MAX_THREAD_LIMIT, STANDARD_EMOJIS } from '../constants.js';
import { ErrorCode } from '../types/errors.js';

// ─────────────────────────────────────────────────────────────────────────────
// Schemas
// ─────────────────────────────────────────────────────────────────────────────

export const SendMessageInputSchema = z.object({
  content: z.string().min(1, 'Message content cannot be empty'),
  conversationId: z.string().optional().default(SELF_CHAT_ID),
  replyToMessageId: z.string().optional(),
});

export const FavoriteInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
});

export const SaveMessageInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  messageId: z.string().min(1, 'Message ID cannot be empty'),
  rootMessageId: z.string().optional(),
});

export const GetChatInputSchema = z.object({
  userId: z.string().min(1, 'User ID cannot be empty'),
});

export const CreateGroupChatInputSchema = z.object({
  userIds: z.array(z.string().min(1)).min(2, 'At least 2 user IDs are required for a group chat'),
  topic: z.string().optional(),
});

export const EditMessageInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  messageId: z.string().min(1, 'Message ID cannot be empty'),
  content: z.string().min(1, 'Content cannot be empty'),
});

export const DeleteMessageInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  messageId: z.string().min(1, 'Message ID cannot be empty'),
});

export const GetUnreadInputSchema = z.object({
  conversationId: z.string().optional(),
});

export const MarkAsReadInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  messageId: z.string().min(1, 'Message ID cannot be empty'),
});

export const GetActivityInputSchema = z.object({
  limit: z.number().min(1).max(200).optional(),
});

export const SearchEmojiInputSchema = z.object({
  query: z.string().min(1, 'Search query cannot be empty'),
});

export const AddReactionInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  messageId: z.string().min(1, 'Message ID cannot be empty'),
  emoji: z.string().min(1, 'Emoji key cannot be empty'),
});

export const RemoveReactionInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  messageId: z.string().min(1, 'Message ID cannot be empty'),
  emoji: z.string().min(1, 'Emoji key cannot be empty'),
});

export const GetSavedMessagesInputSchema = z.object({
  limit: z.number().min(1).max(MAX_THREAD_LIMIT).optional(),
});

export const GetFollowedThreadsInputSchema = z.object({
  limit: z.number().min(1).max(MAX_THREAD_LIMIT).optional(),
});

// ─────────────────────────────────────────────────────────────────────────────
// Tool Definitions
// ─────────────────────────────────────────────────────────────────────────────

const sendMessageToolDefinition: Tool = {
  name: 'teams_send_message',
  description: 'Send a message to a Teams conversation. Use markdown for formatting (not HTML): **bold**, *italic*, ~~strikethrough~~, `code`, ```code blocks```, lists, and newlines. Supports @mentions using @[Name](mri) syntax inline. Example: "Hey @[John Smith](8:orgid:abc...), check this". Get MRI from teams_search_people. Defaults to self-notes (48:notes). For channel thread replies, provide replyToMessageId.',
  inputSchema: {
    type: 'object',
    properties: {
      content: {
        type: 'string',
        description: 'The message content in markdown (not HTML). Supports: **bold**, *italic*, ~~strikethrough~~, `inline code`, ```code blocks```, bullet lists (- item), numbered lists (1. item), and newlines. Do NOT send raw HTML tags. For @mentions, use @[DisplayName](mri) syntax. Example: "Hey @[John Smith](8:orgid:abc...), can you review this?"',
      },
      conversationId: {
        type: 'string',
        description: 'The conversation ID to send to. Use "48:notes" for self-chat (default), or a channel/chat conversation ID.',
      },
      replyToMessageId: {
        type: 'string',
        description: 'For channel thread replies: the message ID of the thread root. Use serverMessageId from teams_send_message, id from teams_get_thread, or messageId from teams_search.',
      },
    },
    required: ['content'],
  },
};

const getFavoritesToolDefinition: Tool = {
  name: 'teams_get_favorites',
  description: 'Get the user\'s favourite/pinned conversations in Teams. Returns conversation IDs with display names (channel name, chat topic, or participant names) and type (Channel, Chat, Meeting).',
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

const addFavoriteToolDefinition: Tool = {
  name: 'teams_add_favorite',
  description: 'Add a conversation to the user\'s favourites/pinned list.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID to pin (e.g., "19:abc@thread.tacv2")',
      },
    },
    required: ['conversationId'],
  },
};

const removeFavoriteToolDefinition: Tool = {
  name: 'teams_remove_favorite',
  description: 'Remove a conversation from the user\'s favourites/pinned list.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID to unpin',
      },
    },
    required: ['conversationId'],
  },
};

const saveMessageToolDefinition: Tool = {
  name: 'teams_save_message',
  description: 'Save (bookmark) a message in Teams. Saved messages can be accessed later from the Saved view in Teams.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID containing the message',
      },
      messageId: {
        type: 'string',
        description: 'The message ID to save. Use: serverMessageId from teams_send_message, id from teams_get_thread, or messageId from teams_search.',
      },
      rootMessageId: {
        type: 'string',
        description: 'For channel threaded replies only: the ID of the thread root post. Not needed for top-level posts or non-channel conversations.',
      },
    },
    required: ['conversationId', 'messageId'],
  },
};

const unsaveMessageToolDefinition: Tool = {
  name: 'teams_unsave_message',
  description: 'Remove a saved (bookmarked) message in Teams.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID containing the message',
      },
      messageId: {
        type: 'string',
        description: 'The message ID to unsave',
      },
      rootMessageId: {
        type: 'string',
        description: 'For channel threaded replies only: the ID of the thread root post. Not needed for top-level posts or non-channel conversations.',
      },
    },
    required: ['conversationId', 'messageId'],
  },
};

const getChatToolDefinition: Tool = {
  name: 'teams_get_chat',
  description: 'Get the conversation ID for a 1:1 chat with a person. Use this to start a new chat or find an existing one. The conversation ID can then be used with teams_send_message to send messages.',
  inputSchema: {
    type: 'object',
    properties: {
      userId: {
        type: 'string',
        description: 'The user\'s identifier. Can be: MRI (8:orgid:guid), object ID with tenant (guid@tenantId), or raw object ID (guid). Get this from teams_search_people results.',
      },
    },
    required: ['userId'],
  },
};

const createGroupChatToolDefinition: Tool = {
  name: 'teams_create_group_chat',
  description: 'Create a new group chat with multiple people. Returns a conversation ID for use with teams_send_message. You are automatically included as a member. For 1:1 chats, use teams_get_chat instead.',
  inputSchema: {
    type: 'object',
    properties: {
      userIds: {
        type: 'array',
        items: { type: 'string' },
        description: 'Array of OTHER users to add (at least 2, not including yourself). Can be: MRI (8:orgid:guid), object ID with tenant (guid@tenantId), or raw GUID. Get these from teams_search_people or teams_get_frequent_contacts.',
        minItems: 2,
      },
      topic: {
        type: 'string',
        description: 'Optional chat name/topic. If omitted, Teams shows member names.',
      },
    },
    required: ['userIds'],
  },
};

const editMessageToolDefinition: Tool = {
  name: 'teams_edit_message',
  description: 'Edit one of your own messages. You can only edit messages you sent. The API will reject attempts to edit other users\' messages.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID containing the message',
      },
      messageId: {
        type: 'string',
        description: 'The message ID to edit. Use: serverMessageId from teams_send_message, id from teams_get_thread, or messageId from teams_search.',
      },
      content: {
        type: 'string',
        description: 'The new content for the message. Can include basic HTML formatting.',
      },
    },
    required: ['conversationId', 'messageId', 'content'],
  },
};

const deleteMessageToolDefinition: Tool = {
  name: 'teams_delete_message',
  description: 'Delete one of your own messages (soft delete - the message remains but content becomes empty). You can only delete messages you sent, unless you are a channel owner/moderator.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID containing the message',
      },
      messageId: {
        type: 'string',
        description: 'The message ID to delete. Use: serverMessageId from teams_send_message, id from teams_get_thread, or messageId from teams_search.',
      },
    },
    required: ['conversationId', 'messageId'],
  },
};

const getUnreadToolDefinition: Tool = {
  name: 'teams_get_unread',
  description: 'Get unread message status. Without parameters, returns aggregate unread counts across all favourite/pinned conversations. With a conversationId, returns unread status for that specific conversation.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'Optional. A specific conversation ID to check. If omitted, checks all favourites.',
      },
    },
  },
};

const markAsReadToolDefinition: Tool = {
  name: 'teams_mark_read',
  description: 'Mark a conversation as read up to a specific message. This updates your read position so messages up to (and including) the specified message are marked as read.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID to mark as read',
      },
      messageId: {
        type: 'string',
        description: 'The message ID to mark as read up to (all messages up to this point will be marked read)',
      },
    },
    required: ['conversationId', 'messageId'],
  },
};

const getActivityToolDefinition: Tool = {
  name: 'teams_get_activity',
  description: 'Get the user\'s activity feed - mentions, reactions, replies, and other notifications. Returns recent activity items with sender, content, and source conversation context.',
  inputSchema: {
    type: 'object',
    properties: {
      limit: {
        type: 'number',
        description: 'Maximum number of activity items to return (default: 50, max: 200)',
      },
    },
  },
};

const searchEmojiToolDefinition: Tool = {
  name: 'teams_search_emoji',
  description: 'Search for emojis by name or keyword. Returns both standard Teams emojis and custom organisation emojis, indicating which is which. Use the returned key with teams_add_reaction.',
  inputSchema: {
    type: 'object',
    properties: {
      query: {
        type: 'string',
        description: 'Search term (e.g., "thumbs", "heart", "laugh", "cat")',
      },
    },
    required: ['query'],
  },
};

const addReactionToolDefinition: Tool = {
  name: 'teams_add_reaction',
  description: 'Add an emoji reaction to a message. Common reactions: like (👍), heart (❤️), laugh (😂), surprised (😮), sad (😢), angry (😠). Use teams_search_emoji to find other emojis.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID containing the message',
      },
      messageId: {
        type: 'string',
        description: 'The message ID to react to. Use: serverMessageId from teams_send_message, id from teams_get_thread, or messageId from teams_search. NOT the messageId from teams_send_message (that is client-generated and will fail).',
      },
      emoji: {
        type: 'string',
        description: 'The emoji key (e.g., "like", "heart", "laugh"). Get from teams_search_emoji or use common ones directly.',
      },
    },
    required: ['conversationId', 'messageId', 'emoji'],
  },
};

const removeReactionToolDefinition: Tool = {
  name: 'teams_remove_reaction',
  description: 'Remove an emoji reaction from a message.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID containing the message',
      },
      messageId: {
        type: 'string',
        description: 'The message ID to remove the reaction from. Use: serverMessageId from teams_send_message, id from teams_get_thread, or messageId from teams_search.',
      },
      emoji: {
        type: 'string',
        description: 'The emoji key to remove (e.g., "like", "heart")',
      },
    },
    required: ['conversationId', 'messageId', 'emoji'],
  },
};

const getSavedMessagesToolDefinition: Tool = {
  name: 'teams_get_saved_messages',
  description: 'Get the list of messages the user has saved (bookmarked) in Teams. Returns references to saved messages with source conversation IDs and direct links. Use teams_get_thread with the sourceConversationId to fetch actual message content.',
  inputSchema: {
    type: 'object',
    properties: {
      limit: {
        type: 'number',
        description: 'Maximum number of saved messages to return (default: 50, max: 200)',
      },
    },
  },
};

const getFollowedThreadsToolDefinition: Tool = {
  name: 'teams_get_followed_threads',
  description: 'Get the list of threads the user is following in Teams. Returns references to followed threads with source conversation IDs and direct links. Use teams_get_thread with the sourceConversationId to fetch actual thread content.',
  inputSchema: {
    type: 'object',
    properties: {
      limit: {
        type: 'number',
        description: 'Maximum number of followed threads to return (default: 50, max: 200)',
      },
    },
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// Handlers
// ─────────────────────────────────────────────────────────────────────────────

async function handleSendMessage(
  input: z.infer<typeof SendMessageInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await sendMessage(input.conversationId, input.content, {
    replyToMessageId: input.replyToMessageId,
  });

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  // The timestamp is the server-assigned ID - use this for reactions, threading, edits, etc.
  const serverMessageId = result.value.timestamp ? String(result.value.timestamp) : undefined;

  const response: Record<string, unknown> = {
    messageId: result.value.messageId,
    timestamp: result.value.timestamp,
    conversationId: input.conversationId,
  };

  // Always include serverMessageId - this is the ID to use for reactions, edits, etc.
  if (serverMessageId) {
    response.serverMessageId = serverMessageId;
  }

  // Include replyToMessageId in response if this was a thread reply
  if (input.replyToMessageId) {
    response.replyToMessageId = input.replyToMessageId;
    response.note = 'Message posted as a reply to the thread. Use serverMessageId (not messageId) for reactions, edits, or threading.';
  } else if (serverMessageId) {
    response.note = 'Use serverMessageId (not messageId) for reactions, edits, or threading.';
  }

  return { success: true, data: response };
}

async function handleGetFavorites(
  _input: Record<string, never>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getFavorites();

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      count: result.value.favorites.length,
      favorites: result.value.favorites,
    },
  };
}

async function handleAddFavorite(
  input: z.infer<typeof FavoriteInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await addFavorite(input.conversationId);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: `Added ${input.conversationId} to favourites`,
    },
  };
}

async function handleRemoveFavorite(
  input: z.infer<typeof FavoriteInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await removeFavorite(input.conversationId);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: `Removed ${input.conversationId} from favourites`,
    },
  };
}

async function handleSaveMessage(
  input: z.infer<typeof SaveMessageInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await saveMessage(input.conversationId, input.messageId, input.rootMessageId);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: 'Message saved',
      conversationId: input.conversationId,
      messageId: input.messageId,
    },
  };
}

async function handleUnsaveMessage(
  input: z.infer<typeof SaveMessageInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await unsaveMessage(input.conversationId, input.messageId, input.rootMessageId);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: 'Message unsaved',
      conversationId: input.conversationId,
      messageId: input.messageId,
    },
  };
}

async function handleGetChat(
  input: z.infer<typeof GetChatInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = getOneOnOneChatId(input.userId);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      conversationId: result.value.conversationId,
      otherUserId: result.value.otherUserId,
      currentUserId: result.value.currentUserId,
      note: 'Use this conversationId with teams_send_message to send a message. The conversation is created automatically when the first message is sent.',
    },
  };
}

async function handleCreateGroupChat(
  input: z.infer<typeof CreateGroupChatInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await createGroupChat(input.userIds, input.topic);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      conversationId: result.value.conversationId,
      members: result.value.members,
      topic: result.value.topic,
      note: 'Use this conversationId with teams_send_message to send messages to the group.',
    },
  };
}

async function handleEditMessage(
  input: z.infer<typeof EditMessageInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await editMessage(
    input.conversationId,
    input.messageId,
    input.content
  );

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: 'Message edited successfully',
      conversationId: result.value.conversationId,
      messageId: result.value.messageId,
    },
  };
}

async function handleDeleteMessage(
  input: z.infer<typeof DeleteMessageInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await deleteMessage(
    input.conversationId,
    input.messageId
  );

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: 'Message deleted successfully',
      conversationId: result.value.conversationId,
      messageId: result.value.messageId,
    },
  };
}

async function handleGetUnread(
  input: z.infer<typeof GetUnreadInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  // If a specific conversation is provided, just check that one
  if (input.conversationId) {
    const result = await getUnreadStatus(input.conversationId);
    if (!result.ok) {
      return { success: false, error: result.error };
    }

    return {
      success: true,
      data: {
        conversationId: result.value.conversationId,
        unreadCount: result.value.unreadCount,
        lastReadMessageId: result.value.lastReadMessageId,
        latestMessageId: result.value.latestMessageId,
      },
    };
  }

  // Aggregate mode: check all favourites
  const favResult = await getFavorites();
  if (!favResult.ok) {
    return { success: false, error: favResult.error };
  }

  const favorites = favResult.value.favorites;
  const conversations: Array<{
    conversationId: string;
    displayName?: string;
    conversationType?: string;
    unreadCount: number;
  }> = [];

  let totalUnread = 0;
  let checkedCount = 0;
  let errorCount = 0;

  // Check unread status for each favourite (limit to prevent timeout)
  const maxToCheck = MAX_UNREAD_AGGREGATE_CHECK;
  for (const fav of favorites.slice(0, maxToCheck)) {
    const unreadResult = await getUnreadStatus(fav.conversationId);
    checkedCount++;

    if (unreadResult.ok) {
      if (unreadResult.value.unreadCount > 0) {
        conversations.push({
          conversationId: fav.conversationId,
          displayName: fav.displayName,
          conversationType: fav.conversationType,
          unreadCount: unreadResult.value.unreadCount,
        });
        totalUnread += unreadResult.value.unreadCount;
      }
    } else {
      errorCount++;
    }
  }

  // If all checks failed, return an error rather than misleading success
  if (checkedCount > 0 && errorCount === checkedCount) {
    return {
      success: false,
      error: {
        code: ErrorCode.API_ERROR,
        message: `Failed to check unread status for all ${checkedCount} favourites`,
        retryable: true,
        suggestions: ['Check authentication status with teams_status', 'Try teams_login to refresh session'],
      },
    };
  }

  return {
    success: true,
    data: {
      totalUnread,
      conversationsWithUnread: conversations.length,
      conversations,
      checked: checkedCount,
      totalFavorites: favorites.length,
      errors: errorCount > 0 ? errorCount : undefined,
      note: favorites.length > maxToCheck
        ? `Checked first ${maxToCheck} of ${favorites.length} favourites`
        : undefined,
    },
  };
}

async function handleMarkAsRead(
  input: z.infer<typeof MarkAsReadInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await markAsRead(input.conversationId, input.messageId);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: 'Conversation marked as read',
      conversationId: result.value.conversationId,
      markedUpTo: result.value.markedUpTo,
    },
  };
}

async function handleGetActivity(
  input: z.infer<typeof GetActivityInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getActivityFeed({ limit: input.limit });

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      count: result.value.activities.length,
      activities: result.value.activities,
      syncState: result.value.syncState,
    },
  };
}

async function handleSearchEmoji(
  input: z.infer<typeof SearchEmojiInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const query = input.query.toLowerCase();
  
  // Search standard emojis
  const standardMatches = STANDARD_EMOJIS.filter(emoji =>
    emoji.key.toLowerCase().includes(query) ||
    emoji.description.toLowerCase().includes(query)
  ).map(emoji => ({
    key: emoji.key,
    description: emoji.description,
    type: 'standard' as const,
    category: emoji.category,
  }));

  // Try to get custom emojis
  let customMatches: Array<{
    key: string;
    description: string;
    type: 'custom';
    shortcut: string;
  }> = [];

  const customResult = await getCustomEmojis();
  if (customResult.ok) {
    customMatches = customResult.value.emojis
      .filter(emoji =>
        emoji.shortcut.toLowerCase().includes(query) ||
        emoji.description.toLowerCase().includes(query)
      )
      .map(emoji => ({
        key: emoji.id,
        description: emoji.description,
        type: 'custom' as const,
        shortcut: emoji.shortcut,
      }));
  }

  // Combine results, standard first
  const results = [...standardMatches, ...customMatches];

  return {
    success: true,
    data: {
      query: input.query,
      count: results.length,
      emojis: results,
      note: results.length === 0
        ? 'No emojis found. Try a different search term.'
        : 'Use the "key" value with teams_add_reaction.',
    },
  };
}

async function handleAddReaction(
  input: z.infer<typeof AddReactionInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await addReaction(
    input.conversationId,
    input.messageId,
    input.emoji
  );

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: `Added ${input.emoji} reaction`,
      conversationId: result.value.conversationId,
      messageId: result.value.messageId,
      emoji: result.value.emoji,
    },
  };
}

async function handleRemoveReaction(
  input: z.infer<typeof RemoveReactionInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await removeReaction(
    input.conversationId,
    input.messageId,
    input.emoji
  );

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: `Removed ${input.emoji} reaction`,
      conversationId: result.value.conversationId,
      messageId: result.value.messageId,
      emoji: result.value.emoji,
    },
  };
}

async function handleGetSavedMessages(
  input: z.infer<typeof GetSavedMessagesInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getSavedMessages({ limit: input.limit });

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      count: result.value.messages.length,
      messages: result.value.messages.map(msg => ({
        content: msg.content,
        contentType: msg.contentType,
        sender: msg.sender,
        timestamp: msg.timestamp,
        sourceConversationId: msg.sourceConversationId,
        sourceMessageId: msg.sourceMessageId,
        messageLink: msg.messageLink,
      })),
    },
  };
}

async function handleGetFollowedThreads(
  input: z.infer<typeof GetFollowedThreadsInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getFollowedThreads({ limit: input.limit });

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      count: result.value.threads.length,
      threads: result.value.threads.map(thread => ({
        content: thread.content,
        contentType: thread.contentType,
        sender: thread.sender,
        timestamp: thread.timestamp,
        sourceConversationId: thread.sourceConversationId,
        sourcePostId: thread.sourcePostId,
        messageLink: thread.messageLink,
      })),
    },
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Exports
// ─────────────────────────────────────────────────────────────────────────────

export const sendMessageTool: RegisteredTool<typeof SendMessageInputSchema> = {
  definition: sendMessageToolDefinition,
  schema: SendMessageInputSchema,
  handler: handleSendMessage,
};

export const getFavoritesTool: RegisteredTool<z.ZodObject<Record<string, never>>> = {
  definition: getFavoritesToolDefinition,
  schema: z.object({}),
  handler: handleGetFavorites,
};

export const addFavoriteTool: RegisteredTool<typeof FavoriteInputSchema> = {
  definition: addFavoriteToolDefinition,
  schema: FavoriteInputSchema,
  handler: handleAddFavorite,
};

export const removeFavoriteTool: RegisteredTool<typeof FavoriteInputSchema> = {
  definition: removeFavoriteToolDefinition,
  schema: FavoriteInputSchema,
  handler: handleRemoveFavorite,
};

export const saveMessageTool: RegisteredTool<typeof SaveMessageInputSchema> = {
  definition: saveMessageToolDefinition,
  schema: SaveMessageInputSchema,
  handler: handleSaveMessage,
};

export const unsaveMessageTool: RegisteredTool<typeof SaveMessageInputSchema> = {
  definition: unsaveMessageToolDefinition,
  schema: SaveMessageInputSchema,
  handler: handleUnsaveMessage,
};

export const getChatTool: RegisteredTool<typeof GetChatInputSchema> = {
  definition: getChatToolDefinition,
  schema: GetChatInputSchema,
  handler: handleGetChat,
};

export const createGroupChatTool: RegisteredTool<typeof CreateGroupChatInputSchema> = {
  definition: createGroupChatToolDefinition,
  schema: CreateGroupChatInputSchema,
  handler: handleCreateGroupChat,
};

export const editMessageTool: RegisteredTool<typeof EditMessageInputSchema> = {
  definition: editMessageToolDefinition,
  schema: EditMessageInputSchema,
  handler: handleEditMessage,
};

export const deleteMessageTool: RegisteredTool<typeof DeleteMessageInputSchema> = {
  definition: deleteMessageToolDefinition,
  schema: DeleteMessageInputSchema,
  handler: handleDeleteMessage,
};

export const getUnreadTool: RegisteredTool<typeof GetUnreadInputSchema> = {
  definition: getUnreadToolDefinition,
  schema: GetUnreadInputSchema,
  handler: handleGetUnread,
};

export const markAsReadTool: RegisteredTool<typeof MarkAsReadInputSchema> = {
  definition: markAsReadToolDefinition,
  schema: MarkAsReadInputSchema,
  handler: handleMarkAsRead,
};

export const getActivityTool: RegisteredTool<typeof GetActivityInputSchema> = {
  definition: getActivityToolDefinition,
  schema: GetActivityInputSchema,
  handler: handleGetActivity,
};

export const searchEmojiTool: RegisteredTool<typeof SearchEmojiInputSchema> = {
  definition: searchEmojiToolDefinition,
  schema: SearchEmojiInputSchema,
  handler: handleSearchEmoji,
};

export const addReactionTool: RegisteredTool<typeof AddReactionInputSchema> = {
  definition: addReactionToolDefinition,
  schema: AddReactionInputSchema,
  handler: handleAddReaction,
};

export const removeReactionTool: RegisteredTool<typeof RemoveReactionInputSchema> = {
  definition: removeReactionToolDefinition,
  schema: RemoveReactionInputSchema,
  handler: handleRemoveReaction,
};

export const getSavedMessagesTool: RegisteredTool<typeof GetSavedMessagesInputSchema> = {
  definition: getSavedMessagesToolDefinition,
  schema: GetSavedMessagesInputSchema,
  handler: handleGetSavedMessages,
};

export const getFollowedThreadsTool: RegisteredTool<typeof GetFollowedThreadsInputSchema> = {
  definition: getFollowedThreadsToolDefinition,
  schema: GetFollowedThreadsInputSchema,
  handler: handleGetFollowedThreads,
};

/** All message-related tools. */
export const messageTools = [
  sendMessageTool,
  getFavoritesTool,
  addFavoriteTool,
  removeFavoriteTool,
  saveMessageTool,
  unsaveMessageTool,
  getChatTool,
  createGroupChatTool,
  editMessageTool,
  deleteMessageTool,
  getUnreadTool,
  markAsReadTool,
  getActivityTool,
  searchEmojiTool,
  addReactionTool,
  removeReactionTool,
  getSavedMessagesTool,
  getFollowedThreadsTool,
];
