/**
 * Microsoft Graph API - Spike for sending messages.
 * 
 * This is a spike/experiment to test whether we can use the Graph API
 * with tokens extracted from the Teams browser session (no Azure App
 * registration required).
 * 
 * Graph API send message endpoint:
 * POST https://graph.microsoft.com/v1.0/chats/{chatId}/messages
 * 
 * Note: Graph API uses a different chat ID format than chatsvc. The
 * conversationId from Teams (e.g., "19:xxx@thread.v2") should work
 * directly as the Graph chatId.
 */

import { httpRequest } from '../utils/http.js';
import { ErrorCode, createError } from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';
import { requireGraphAuth } from '../utils/auth-guards.js';

// ─────────────────────────────────────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────────────────────────────────────

const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/** Result of sending a message via Graph API. */
export interface GraphSendMessageResult {
  /** The Graph-assigned message ID. */
  id: string;
  /** When the message was created (ISO 8601). */
  createdDateTime?: string;
  /** The web URL for the message (if returned). */
  webUrl?: string;
  /** Raw Graph API response for debugging in spike. */
  _raw?: unknown;
}

/** Graph API error response shape. */
interface GraphErrorResponse {
  error?: {
    code?: string;
    message?: string;
    innerError?: {
      code?: string;
      message?: string;
      date?: string;
      'request-id'?: string;
    };
  };
}

/** Graph API message response shape. */
interface GraphMessageResponse {
  id?: string;
  createdDateTime?: string;
  webUrl?: string;
  body?: {
    contentType?: string;
    content?: string;
  };
  from?: {
    user?: {
      displayName?: string;
      id?: string;
    };
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Send Message via Graph API
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Sends a message to a Teams chat via the Microsoft Graph API.
 * 
 * This is a spike to test Graph API access using tokens from the Teams
 * browser session. The Graph API endpoint is:
 * POST /v1.0/chats/{chatId}/messages
 * 
 * For channel messages, the endpoint would be:
 * POST /v1.0/teams/{teamId}/channels/{channelId}/messages
 * 
 * @param chatId - The Teams conversation ID (e.g., "19:xxx@thread.v2")
 * @param content - The message content (plain text or HTML)
 * @param contentType - "text" for plain text, "html" for HTML (default: "text")
 */
export async function graphSendMessage(
  chatId: string,
  content: string,
  contentType: 'text' | 'html' = 'text'
): Promise<Result<GraphSendMessageResult>> {
  const authResult = requireGraphAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const { graphToken } = authResult.value;

  const url = `${GRAPH_BASE_URL}/chats/${encodeURIComponent(chatId)}/messages`;

  const body = {
    body: {
      contentType,
      content,
    },
  };

  const response = await httpRequest<GraphMessageResponse | GraphErrorResponse>(
    url,
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${graphToken}`,
      },
      body: JSON.stringify(body),
      // Don't retry auth errors - they indicate the token doesn't have the right scopes
      maxRetries: 1,
    }
  );

  if (!response.ok) {
    // Enhance error message with Graph-specific context
    const errorMessage = response.error.message;
    
    // Check for common Graph API permission errors
    if (errorMessage.includes('Authorization_RequestDenied') || 
        errorMessage.includes('Forbidden') ||
        errorMessage.includes('403')) {
      return err(createError(
        ErrorCode.AUTH_REQUIRED,
        `Graph API permission denied. The token from the Teams session may not have Chat.ReadWrite scope. Error: ${errorMessage}`,
        { 
          retryable: false,
          suggestions: [
            'The Teams SPA client ID may not have delegated Chat.ReadWrite permission',
            'Check the token scopes by decoding the JWT at jwt.ms',
          ],
        }
      ));
    }

    return response;
  }

  const data = response.value.data;

  // Check if the response is an error response
  if ('error' in data && data.error) {
    return err(createError(
      ErrorCode.API_ERROR,
      `Graph API error: ${data.error.code}: ${data.error.message}`,
      { retryable: false }
    ));
  }

  const msgData = data as GraphMessageResponse;

  if (!msgData.id) {
    return err(createError(
      ErrorCode.UNKNOWN,
      'Graph API returned success but no message ID',
      { retryable: false }
    ));
  }

  return ok({
    id: msgData.id,
    createdDateTime: msgData.createdDateTime,
    webUrl: msgData.webUrl,
    _raw: msgData,
  });
}

/**
 * Sends a message to a Teams channel via the Microsoft Graph API.
 * 
 * POST /v1.0/teams/{teamId}/channels/{channelId}/messages
 * 
 * Note: For channels, we need the teamId (group ID) and channelId separately.
 * This is different from the chatsvc API which uses a single conversationId.
 * 
 * @param teamId - The team's group ID (GUID)
 * @param channelId - The channel ID (e.g., "19:xxx@thread.tacv2")
 * @param content - The message content
 * @param contentType - "text" for plain text, "html" for HTML (default: "text")
 */
export async function graphSendChannelMessage(
  teamId: string,
  channelId: string,
  content: string,
  contentType: 'text' | 'html' = 'text'
): Promise<Result<GraphSendMessageResult>> {
  const authResult = requireGraphAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const { graphToken } = authResult.value;

  const url = `${GRAPH_BASE_URL}/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages`;

  const body = {
    body: {
      contentType,
      content,
    },
  };

  const response = await httpRequest<GraphMessageResponse | GraphErrorResponse>(
    url,
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${graphToken}`,
      },
      body: JSON.stringify(body),
      maxRetries: 1,
    }
  );

  if (!response.ok) {
    return response;
  }

  const data = response.value.data;

  if ('error' in data && data.error) {
    return err(createError(
      ErrorCode.API_ERROR,
      `Graph API error: ${data.error.code}: ${data.error.message}`,
      { retryable: false }
    ));
  }

  const msgData = data as GraphMessageResponse;

  if (!msgData.id) {
    return err(createError(
      ErrorCode.UNKNOWN,
      'Graph API returned success but no message ID',
      { retryable: false }
    ));
  }

  return ok({
    id: msgData.id,
    createdDateTime: msgData.createdDateTime,
    webUrl: msgData.webUrl,
    _raw: msgData,
  });
}
