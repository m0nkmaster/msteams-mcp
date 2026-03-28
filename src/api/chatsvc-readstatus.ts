/**
 * Chat Service API - Read status operations.
 * 
 * Consumption horizons (read receipts), mark as read, and unread counts.
 */

import { httpRequest } from '../utils/http.js';
import { CHATSVC_API, getSkypeAuthHeaders } from '../utils/api-config.js';
import { type Result, ok } from '../types/result.js';
import { requireMessageAuthWithConfig } from '../utils/auth-guards.js';
import { getThreadMessages } from './chatsvc-messaging.js';
import type { RawConsumptionHorizonsResponse } from '../types/api-responses.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/** Consumption horizon information for a conversation. */
export interface ConsumptionHorizonInfo {
  /** The conversation/thread ID. */
  conversationId: string;
  /** The version timestamp of the consumption horizon. */
  version?: string;
  /** The last read message ID (timestamp). */
  lastReadMessageId?: string;
  /** The last read timestamp. */
  lastReadTimestamp?: number;
  /** Raw consumption horizons array from API. */
  consumptionHorizons: Array<{
    id: string;
    consumptionHorizon: string;
  }>;
}

/** Result of marking a conversation as read. */
export interface MarkAsReadResult {
  conversationId: string;
  markedUpTo: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// API Functions
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Gets the consumption horizon (read receipts) for a conversation.
 * The consumption horizon indicates where each user has read up to.
 */
export async function getConsumptionHorizon(
  conversationId: string
): Promise<Result<ConsumptionHorizonInfo>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, region, baseUrl } = authResult.value;
  const url = CHATSVC_API.consumptionHorizons(region, conversationId, baseUrl);

  const response = await httpRequest<RawConsumptionHorizonsResponse>(
    url,
    {
      method: 'GET',
      headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken, baseUrl),
    }
  );

  if (!response.ok) {
    return response;
  }

  const data = response.value.data;
  const horizons = data.consumptionhorizons || [];

  // Find the current user's consumption horizon
  let lastReadMessageId: string | undefined;
  let lastReadTimestamp: number | undefined;

  for (const h of horizons) {
    if (h.id === auth.userMri || h.id.includes(auth.userMri)) {
      // Consumption horizon format: "{timestamp};{timestamp};{messageId}"
      const parts = h.consumptionhorizon.split(';');
      if (parts.length >= 3) {
        lastReadMessageId = parts[2];
        lastReadTimestamp = parseInt(parts[0], 10);
      }
      break;
    }
  }

  return ok({
    conversationId,
    version: data.version,
    lastReadMessageId,
    lastReadTimestamp,
    consumptionHorizons: horizons.map(h => ({
      id: h.id,
      consumptionHorizon: h.consumptionhorizon,
    })),
  });
}

/**
 * Marks a conversation as read up to a specific message.
 */
export async function markAsRead(
  conversationId: string,
  messageId: string
): Promise<Result<MarkAsReadResult>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, region, baseUrl } = authResult.value;
  const url = CHATSVC_API.updateConsumptionHorizon(region, conversationId, baseUrl);

  // Format: "{messageId};{messageId};{messageId}" - all three values are the same
  const consumptionHorizon = `${messageId};${messageId};${messageId}`;

  const response = await httpRequest<unknown>(
    url,
    {
      method: 'PUT',
      headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken, baseUrl),
      body: JSON.stringify({
        consumptionhorizon: consumptionHorizon,
      }),
    }
  );

  if (!response.ok) {
    return response;
  }

  return ok({
    conversationId,
    markedUpTo: messageId,
  });
}

/**
 * Gets unread count for a conversation by comparing consumption horizon
 * with recent messages.
 *
 * If the consumption horizon endpoint fails, falls back to message-only
 * checking (reports recent messages from others without precise read position).
 */
export async function getUnreadStatus(
  conversationId: string
): Promise<Result<{
  conversationId: string;
  unreadCount: number;
  lastReadMessageId?: string;
  latestMessageId?: string;
}>> {
  // Get consumption horizon — non-fatal if it fails
  const horizonResult = await getConsumptionHorizon(conversationId);
  const lastReadId = horizonResult.ok ? horizonResult.value.lastReadMessageId : undefined;

  // Get recent messages
  const messagesResult = await getThreadMessages(conversationId, { limit: 50 });
  if (!messagesResult.ok) {
    return messagesResult;
  }

  const messages = messagesResult.value.messages;

  // Count messages after the last read position
  let unreadCount = 0;
  let latestMessageId: string | undefined;

  // Messages from getThreadMessages are newest-first by default (desc order).
  // Iterate from newest to oldest to count unread messages after lastReadId.
  for (const msg of messages) {
    if (!latestMessageId && !msg.isFromMe) {
      latestMessageId = msg.id;
    }

    if (lastReadId && msg.id === lastReadId) {
      break;
    }

    // Count messages not from the current user
    if (!msg.isFromMe) {
      unreadCount++;
    }
  }

  // If last read message wasn't in our window, count is a lower bound

  return ok({
    conversationId,
    unreadCount,
    lastReadMessageId: lastReadId,
    latestMessageId,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Bulk unread check via conversations list
// ─────────────────────────────────────────────────────────────────────────────

/** An unread conversation from the bulk check. */
export interface UnreadConversation {
  conversationId: string;
  displayName?: string;
  isChannel: boolean;
  lastMessageFrom?: string;
  lastMessageTime: number;
  readUpTo: number;
}

/** Result of the bulk unread conversations check. */
export interface UnreadConversationsResult {
  unreadChats: UnreadConversation[];
  unreadChannels: UnreadConversation[];
  totalChecked: number;
}

/**
 * Gets all unread conversations in a single API call.
 *
 * Uses the conversations list endpoint which returns each conversation's
 * consumptionhorizon inline, avoiding N+1 API calls. Compares the last
 * message timestamp against the read horizon to determine unread state.
 */
export async function getUnreadConversations(): Promise<Result<UnreadConversationsResult>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, region, baseUrl } = authResult.value;

  const url = CHATSVC_API.conversations(region, baseUrl);
  const response = await httpRequest<Record<string, unknown>>(
    url,
    {
      method: 'GET',
      headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken, baseUrl),
    }
  );

  if (!response.ok) {
    return response;
  }

  const data = response.value.data as Record<string, unknown> | undefined;
  const convs = (data?.conversations as unknown[]) || [];

  const unreadChats: UnreadConversation[] = [];
  const unreadChannels: UnreadConversation[] = [];

  for (const raw of convs) {
    const c = raw as Record<string, unknown>;
    const props = (c.properties || {}) as Record<string, string>;
    const tp = (c.threadProperties || {}) as Record<string, string>;
    const lastMsg = c.lastMessage as Record<string, string> | undefined;

    if (!lastMsg?.id) continue;

    const lastMsgTime = parseInt(lastMsg.id, 10);
    if (isNaN(lastMsgTime)) continue;
    const fromMe = lastMsg.from?.includes(auth.userMri);
    const isChannel = tp.threadType === 'channel' || (c.id as string).includes('@thread.tacv2');
    const displayName = tp.topic || lastMsg.imdisplayname;

    const horizon = props.consumptionhorizon;

    if (!horizon) {
      // Never-opened conversation — only flag DMs where last message isn't from us
      if (fromMe || isChannel) continue;
      unreadChats.push({
        conversationId: c.id as string,
        displayName,
        isChannel: false,
        lastMessageFrom: lastMsg.imdisplayname,
        lastMessageTime: lastMsgTime,
        readUpTo: 0,
      });
      continue;
    }

    const readUpTo = parseInt(horizon.split(';')[0], 10);
    if (isNaN(readUpTo)) continue; // malformed horizon — skip rather than misclassify
    if (lastMsgTime <= readUpTo) continue; // Already read

    // For channels, preserve original behavior (skip if last msg is ours).
    // For chats, only skip if read horizon is within 2s of our reply —
    // a larger gap means unread messages exist before our reply.
    // 2-second (2000 ms) window: if the read horizon is this close to our last message,
    // assume we've read everything and there's no unread gap before our reply.
    const SELF_REPLY_READ_WINDOW_MS = 2000;
    if (fromMe && (isChannel || (lastMsgTime - readUpTo) < SELF_REPLY_READ_WINDOW_MS)) continue;

    const entry: UnreadConversation = {
      conversationId: c.id as string,
      displayName,
      isChannel,
      lastMessageFrom: lastMsg.imdisplayname,
      lastMessageTime: lastMsgTime,
      readUpTo,
    };

    if (isChannel) {
      unreadChannels.push(entry);
    } else {
      unreadChats.push(entry);
    }
  }

  return ok({
    unreadChats,
    unreadChannels,
    totalChecked: convs.length,
  });
}
