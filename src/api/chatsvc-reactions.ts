/**
 * Chat Service API - Reaction (emoji) operations.
 * 
 * Add and remove emoji reactions on messages.
 */

import { httpRequest } from '../utils/http.js';
import { CHATSVC_API, getSkypeAuthHeaders } from '../utils/api-config.js';
import { type Result, ok } from '../types/result.js';
import { requireMessageAuthWithConfig } from '../utils/auth-guards.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/** Result of adding/removing a reaction. */
export interface ReactionResult {
  conversationId: string;
  messageId: string;
  emoji: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// API Functions
// ─────────────────────────────────────────────────────────────────────────────

/** Shared implementation for adding/removing reactions. */
async function setReaction(
  conversationId: string,
  messageId: string,
  emojiKey: string,
  method: 'PUT' | 'DELETE'
): Promise<Result<ReactionResult>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, region, baseUrl } = authResult.value;
  const url = CHATSVC_API.messageEmotions(region, conversationId, messageId, baseUrl);

  const emotions: Record<string, unknown> = { key: emojiKey };
  if (method === 'PUT') {
    emotions.value = Date.now();
  }

  const response = await httpRequest<unknown>(
    url,
    {
      method,
      headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken, baseUrl),
      body: JSON.stringify({ emotions }),
    }
  );

  if (!response.ok) {
    return response;
  }

  return ok({
    conversationId,
    messageId,
    emoji: emojiKey,
  });
}

/**
 * Adds a reaction (emoji) to a message.
 * 
 * @param conversationId - The conversation containing the message
 * @param messageId - The message ID to react to
 * @param emojiKey - The emoji key (e.g., 'like', 'heart', 'laugh')
 */
export async function addReaction(
  conversationId: string,
  messageId: string,
  emojiKey: string
): Promise<Result<ReactionResult>> {
  return setReaction(conversationId, messageId, emojiKey, 'PUT');
}

/**
 * Removes a reaction (emoji) from a message.
 * 
 * @param conversationId - The conversation containing the message
 * @param messageId - The message ID to remove the reaction from
 * @param emojiKey - The emoji key to remove (e.g., 'like', 'heart')
 */
export async function removeReaction(
  conversationId: string,
  messageId: string,
  emojiKey: string
): Promise<Result<ReactionResult>> {
  return setReaction(conversationId, messageId, emojiKey, 'DELETE');
}
