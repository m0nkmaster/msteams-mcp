/**
 * Chat Service API - Virtual conversation operations.
 * 
 * Saved messages (48:saved) and followed threads (48:threads).
 * Also includes save/unsave message operations.
 */

import { httpRequest } from '../utils/http.js';
import { CHATSVC_API, getSkypeAuthHeaders } from '../utils/api-config.js';
import { type Result, ok } from '../types/result.js';
import { requireMessageAuthWithConfig, getTeamsBaseUrl, getTenantId } from '../utils/auth-guards.js';
import { parseVirtualConversationMessage, type ExtractedLink } from '../utils/parsers.js';
import { SAVED_MESSAGES_ID, FOLLOWED_THREADS_ID } from '../constants.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/** Result of saving/unsaving a message. */
export interface SaveMessageResult {
  conversationId: string;
  messageId: string;
  saved: boolean;
}

/** A saved message from the virtual 48:saved conversation. */
export interface SavedMessage {
  id: string;
  content: string;
  contentType: string;
  sender: {
    mri: string;
    displayName?: string;
  };
  timestamp: string;
  /** The original conversation where this message lives. */
  sourceConversationId: string;
  /** The original message ID in the source conversation. */
  sourceMessageId?: string;
  messageLink?: string;
  links?: ExtractedLink[];
}

/** Result of getting saved messages. */
export interface GetSavedMessagesResult {
  messages: SavedMessage[];
}

/** A followed thread from the virtual 48:threads conversation. */
export interface FollowedThread {
  id: string;
  content: string;
  contentType: string;
  sender: {
    mri: string;
    displayName?: string;
  };
  timestamp: string;
  /** The original conversation/thread where this post lives. */
  sourceConversationId: string;
  /** The original post ID in the source conversation. */
  sourcePostId?: string;
  messageLink?: string;
  links?: ExtractedLink[];
}

/** Result of getting followed threads. */
export interface GetFollowedThreadsResult {
  threads: FollowedThread[];
}

// ─────────────────────────────────────────────────────────────────────────────
// Internal Types & Constants
// ─────────────────────────────────────────────────────────────────────────────

// Regex pattern for extracting message ID from saved messages: T_{conversationId}_M_{messageId}
const SAVED_MESSAGE_PATTERN = /_M_(\d+)$/;

// Regex pattern for extracting post ID from followed threads: T_{conversationId}_P_{postId}_Threads
const FOLLOWED_THREAD_PATTERN = /_P_(\d+)_Threads$/;

/** Parsed item from a virtual conversation (saved messages or followed threads). */
interface VirtualConversationItem {
  id: string;
  content: string;
  contentType: string;
  sender: { mri: string; displayName?: string };
  timestamp: string;
  sourceConversationId: string;
  sourceReferenceId?: string;
  messageLink?: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// Internal Helpers
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Fetches and parses messages from a virtual conversation (48:saved or 48:threads).
 * 
 * Both saved messages and followed threads use the same API pattern:
 * fetch from a virtual conversation ID, parse with a regex pattern, and sort by timestamp.
 */
async function fetchVirtualConversation(
  virtualConversationId: string,
  referencePattern: RegExp,
  options: { limit?: number } = {}
): Promise<Result<VirtualConversationItem[]>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, region, baseUrl } = authResult.value;
  const limit = options.limit ?? 50;

  let url = CHATSVC_API.messages(region, virtualConversationId, undefined, baseUrl);
  url += `?view=msnp24Equivalent|supportsMessageProperties&pageSize=${limit}&startTime=1`;

  const response = await httpRequest<{ messages?: unknown[] }>(
    url,
    {
      method: 'GET',
      headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken, baseUrl),
    }
  );

  if (!response.ok) {
    return response;
  }

  const rawMessages = response.value.data.messages;
  if (!Array.isArray(rawMessages)) {
    return ok([]);
  }

  const linkContext = { tenantId: getTenantId() ?? undefined, teamsBaseUrl: getTeamsBaseUrl() };
  const items: VirtualConversationItem[] = [];

  for (const raw of rawMessages) {
    const parsed = parseVirtualConversationMessage(
      raw as Record<string, unknown>,
      referencePattern,
      linkContext
    );
    if (!parsed) continue;

    items.push({
      id: parsed.id,
      content: parsed.content,
      contentType: parsed.contentType,
      sender: parsed.sender,
      timestamp: parsed.timestamp,
      sourceConversationId: parsed.sourceConversationId,
      sourceReferenceId: parsed.sourceReferenceId,
      messageLink: parsed.messageLink,
    });
  }

  // Sort by timestamp (newest first)
  items.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());

  return ok(items);
}

// ─────────────────────────────────────────────────────────────────────────────
// API Functions
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Gets saved (bookmarked) messages from the virtual 48:saved conversation.
 * Returns messages the user has bookmarked across all conversations.
 */
export async function getSavedMessages(
  options: { limit?: number } = {}
): Promise<Result<GetSavedMessagesResult>> {
  const result = await fetchVirtualConversation(SAVED_MESSAGES_ID, SAVED_MESSAGE_PATTERN, options);
  if (!result.ok) return result;

  return ok({
    messages: result.value.map(item => ({
      ...item,
      sourceMessageId: item.sourceReferenceId,
    })),
  });
}

/**
 * Gets followed threads from the virtual 48:threads conversation.
 * Returns threads the user is following for updates.
 */
export async function getFollowedThreads(
  options: { limit?: number } = {}
): Promise<Result<GetFollowedThreadsResult>> {
  const result = await fetchVirtualConversation(FOLLOWED_THREADS_ID, FOLLOWED_THREAD_PATTERN, options);
  if (!result.ok) return result;

  return ok({
    threads: result.value.map(item => ({
      ...item,
      sourcePostId: item.sourceReferenceId,
    })),
  });
}

/**
 * Saves (bookmarks) a message.
 * 
 * @param rootMessageId - For channel threaded replies, the ID of the thread root post.
 *                        For top-level posts and non-channel messages, omit or pass undefined.
 */
export async function saveMessage(
  conversationId: string,
  messageId: string,
  rootMessageId?: string
): Promise<Result<SaveMessageResult>> {
  return setMessageSavedState(conversationId, messageId, true, rootMessageId);
}

/**
 * Unsaves (removes bookmark from) a message.
 * 
 * @param rootMessageId - For channel threaded replies, the ID of the thread root post.
 *                        For top-level posts and non-channel messages, omit or pass undefined.
 */
export async function unsaveMessage(
  conversationId: string,
  messageId: string,
  rootMessageId?: string
): Promise<Result<SaveMessageResult>> {
  return setMessageSavedState(conversationId, messageId, false, rootMessageId);
}

/**
 * Internal function to set the saved state of a message.
 * 
 * The rcmetadata API uses a two-ID system:
 * - URL path: rootMessageId (thread root for channel replies, or messageId for top-level)
 * - Body: mid (the actual message being saved/unsaved)
 */
async function setMessageSavedState(
  conversationId: string,
  messageId: string,
  saved: boolean,
  rootMessageId: string | undefined
): Promise<Result<SaveMessageResult>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, region, baseUrl } = authResult.value;
  // For channel threaded replies, rootMessageId is the thread root post ID
  // For top-level posts and non-channel messages, use the messageId itself
  const urlMessageId = rootMessageId ?? messageId;
  const url = CHATSVC_API.messageMetadata(region, conversationId, urlMessageId, baseUrl);

  const response = await httpRequest<unknown>(
    url,
    {
      method: 'PUT',
      headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken, baseUrl),
      body: JSON.stringify({
        s: saved ? 1 : 0,
        mid: parseInt(messageId, 10),
      }),
    }
  );

  if (!response.ok) {
    return response;
  }

  return ok({
    conversationId,
    messageId,
    saved,
  });
}
