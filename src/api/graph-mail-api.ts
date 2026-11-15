/**
 * Single-email reads via Microsoft Graph /me/messages.
 *
 * `teams_search_email` results carry the message's REST id (ItemRestId), which
 * Graph accepts directly. Older results may hold the EWS-format Substrate `Id`;
 * when a direct GET of one fails it is translated once via translateExchangeIds
 * (ewsId → restId) and retried.
 *
 * Auth: a Graph token with a Mail.* scope, obtained on demand via
 * `requireGraphMailTokenAsync()` (Teams' own cached Graph token may lack one). Like file downloads it is optional: a failure
 * never surfaces as AUTH_REQUIRED/AUTH_EXPIRED, so it can't trigger the
 * server's auto-login.
 */

import { httpRequest } from '../utils/http.js';
import { type Result, ok, err } from '../types/result.js';
import { type McpError, ErrorCode, createError } from '../types/errors.js';
import type { RawGraphMessage, RawGraphTranslatedIds } from '../types/api-responses.js';
import type { EmailDetail } from '../types/teams.js';
import { requireGraphMailTokenAsync } from '../utils/auth-guards.js';
import { graphAuthError } from './graph-files-api.js';
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

  const tokenResult = await requireGraphMailTokenAsync();
  if (!tokenResult.ok) return tokenResult;
  const token = tokenResult.value;

  let response = await fetchMessage(id, token);
  if (!response.ok && isEwsId(id)
    && (response.error.code === ErrorCode.INVALID_INPUT || response.error.code === ErrorCode.NOT_FOUND)) {
    // Older search results hold EWS-format ids: translate and retry once.
    const translated = await translateToRestId(id, token);
    if (!translated.ok) return err(mapEmailError(translated.error, token, id));
    if (translated.value && translated.value !== id) response = await fetchMessage(translated.value, token);
  }
  if (!response.ok) return err(mapEmailError(response.error, token, id));

  return ok(parseGraphEmail(response.value.data, options.format ?? 'text'));
}

/**
 * EWS ids are standard base64 (`+`, `/`); Graph REST ids are base64url. An id
 * with neither is identical in both formats, so translating it is pointless.
 */
function isEwsId(id: string): boolean {
  return /[+/]/.test(id);
}

function fetchMessage(id: string, token: string) {
  const params = new URLSearchParams({ $select: MESSAGE_SELECT, $expand: ATTACHMENTS_EXPAND });
  return httpRequest<RawGraphMessage>(`${GRAPH_API.message(id)}?${params.toString()}`, {
    method: 'GET',
    headers: graphHeaders(token),
  });
}

/** Translates an EWS-format id to a Graph REST id (undefined when Graph returns none). */
async function translateToRestId(id: string, token: string): Promise<Result<string | undefined>> {
  const response = await httpRequest<RawGraphTranslatedIds>(GRAPH_API.translateExchangeIds, {
    method: 'POST',
    headers: { ...graphHeaders(token), 'Content-Type': 'application/json' },
    body: JSON.stringify({
      inputIds: [id],
      sourceIdType: 'ewsId',
      targetIdType: 'restId',
    }),
  });
  if (!response.ok) return response;
  return ok(response.value.data.value?.[0]?.targetId);
}

/**
 * Auth failures go through the shared Graph mapping (never auto-login), a 404
 * becomes an actionable NOT_FOUND that keeps Graph's message, and anything else
 * (400s, throttling, 5xx) passes through with its own message and retryable flag.
 */
function mapEmailError(error: McpError, token: string, id: string): McpError {
  const authError = graphAuthError(error, token, 'email');
  if (authError) return authError;
  if (error.code !== ErrorCode.NOT_FOUND) return error;
  return createError(ErrorCode.NOT_FOUND,
    `Email not found. The id may be stale or belong to a different mailbox. (${error.message})`,
    {
      retryable: false,
      suggestions: ["Use an 'id' value from teams_search_email results", `Verify the id: ${id.slice(0, 40)}…`],
    });
}
