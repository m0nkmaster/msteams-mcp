/**
 * Single-email reads via Microsoft Graph /me/messages.
 *
 * `teams_search_email` results carry the message's REST id (ItemRestId), which
 * Graph accepts directly. Older results may hold the EWS-format Substrate `Id`;
 * when a direct GET fails those are translated once via translateExchangeIds
 * (ewsId → restId) and retried.
 *
 * Auth: the Teams client's own Graph token (Mail.ReadWrite), obtained on demand
 * via `requireGraphTokenAsync()`. Like file downloads it is optional: a failure
 * never surfaces as AUTH_REQUIRED/AUTH_EXPIRED, so it can't trigger the
 * server's auto-login.
 */

import { httpRequest } from '../utils/http.js';
import { type Result, ok, err } from '../types/result.js';
import { type McpError, ErrorCode, createError } from '../types/errors.js';
import type { RawGraphMessage, RawGraphTranslatedIds } from '../types/api-responses.js';
import type { EmailDetail } from '../types/teams.js';
import { requireGraphTokenAsync } from '../utils/auth-guards.js';
import { invalidateAccessToken } from '../auth/token-extractor.js';
import { parseGraphEmail } from '../utils/parsers.js';
import { GRAPH_API } from '../utils/api-config.js';

const MESSAGE_SELECT =
  'id,subject,from,sender,toRecipients,ccRecipients,bccRecipients,sentDateTime,receivedDateTime,' +
  'importance,isRead,hasAttachments,webLink,conversationId,internetMessageId,body';
const ATTACHMENTS_EXPAND = 'attachments($select=name,contentType,size,isInline)';

function graphHeaders(token: string): Record<string, string> {
  return { Authorization: `Bearer ${token}`, Accept: 'application/json' };
}

/**
 * Fetches one email by id. `id` is an `id` value from teams_search_email
 * results (a Graph REST id; legacy EWS ids are translated automatically).
 * Read-only.
 */
export async function getEmail(
  id: string,
  options: { format?: 'text' | 'html' } = {},
): Promise<Result<EmailDetail>> {
  if (!id || /\s/.test(id)) {
    return err(createError(ErrorCode.INVALID_INPUT,
      'id must be a non-empty message id with no whitespace',
      { suggestions: ["Use an 'id' value from teams_search_email results"] }));
  }

  const tokenResult = await requireGraphTokenAsync();
  if (!tokenResult.ok) return tokenResult;
  const token = tokenResult.value;

  let response = await fetchMessage(id, token);
  if (!response.ok) {
    const authError = handleGraphAuthError(response.error, token);
    if (authError) return authError;

    if (response.error.code === ErrorCode.INVALID_INPUT || response.error.code === ErrorCode.NOT_FOUND) {
      // The id may be EWS-format (older search results): translate and retry once.
      const restId = await translateToRestId(id, token);
      if (restId && restId !== id) {
        response = await fetchMessage(restId, token);
        if (response.ok) return ok(parseGraphEmail(response.value.data, options.format ?? 'text'));
        const retryAuthError = handleGraphAuthError(response.error, token);
        if (retryAuthError) return retryAuthError;
      }
      return err(emailNotFoundError(id));
    }
    return response;
  }

  return ok(parseGraphEmail(response.value.data, options.format ?? 'text'));
}

function fetchMessage(id: string, token: string) {
  const params = new URLSearchParams({ $select: MESSAGE_SELECT, $expand: ATTACHMENTS_EXPAND });
  return httpRequest<RawGraphMessage>(`${GRAPH_API.message(id)}?${params.toString()}`, {
    method: 'GET',
    headers: graphHeaders(token),
  });
}

/** Translates an EWS-format id to a Graph REST id; undefined when untranslatable. */
async function translateToRestId(id: string, token: string): Promise<string | undefined> {
  const response = await httpRequest<RawGraphTranslatedIds>(GRAPH_API.translateExchangeIds, {
    method: 'POST',
    headers: { ...graphHeaders(token), 'Content-Type': 'application/json' },
    body: JSON.stringify({
      inputIds: [id],
      sourceIdType: 'ewsId',
      targetIdType: 'restId',
    }),
  });
  if (!response.ok) return undefined;
  return response.value.data.value?.[0]?.targetId;
}

/**
 * Maps a Graph auth failure the same way file downloads do: a rejected token is
 * discarded and reported as retryable API_ERROR, a forbidden token as
 * ACCESS_DENIED — never AUTH_REQUIRED/AUTH_EXPIRED, which would trigger
 * auto-login. Returns null for non-auth errors.
 */
function handleGraphAuthError(error: McpError, token: string): Result<EmailDetail> | null {
  if (error.code === ErrorCode.AUTH_EXPIRED) {
    invalidateAccessToken(token);
    return err(createError(ErrorCode.API_ERROR,
      'Microsoft Graph rejected its access token; it has been discarded, so a retry will request a new one.',
      { retryable: true }));
  }
  if (error.code === ErrorCode.AUTH_REQUIRED) {
    return err(createError(ErrorCode.ACCESS_DENIED,
      `You don't have access to this email: ${error.message}`, { retryable: false }));
  }
  return null;
}

function emailNotFoundError(id: string): McpError {
  return createError(ErrorCode.NOT_FOUND,
    'Email not found. The id may be stale or belong to a different mailbox.',
    {
      retryable: false,
      suggestions: ["Use an 'id' value from teams_search_email results", `Verify the id: ${id.slice(0, 40)}…`],
    });
}
