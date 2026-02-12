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
 */
export async function getUnreadStatus(
  conversationId: string
): Promise<Result<{
  conversationId: string;
  unreadCount: number;
  lastReadMessageId?: string;
  latestMessageId?: string;
}>> {
  // Get consumption horizon
  const horizonResult = await getConsumptionHorizon(conversationId);
  if (!horizonResult.ok) {
    return horizonResult;
  }

  // Get recent messages
  const messagesResult = await getThreadMessages(conversationId, { limit: 50 });
  if (!messagesResult.ok) {
    return messagesResult;
  }

  const lastReadId = horizonResult.value.lastReadMessageId;
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
