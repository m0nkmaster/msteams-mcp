/**
 * Single-email reads via Microsoft Graph /me/messages.
 *
 * `teams_search_email` results carry the message's Graph REST id, which this
 * fetches directly. Auth is the Teams client's own cached Graph token, which
 * includes Mail.ReadWrite. Like file downloads it is optional: a failure never
 * surfaces as AUTH_REQUIRED/AUTH_EXPIRED, so it can't trigger the server's
 * auto-login.
 */

import { httpRequest } from '../utils/http.js';
import { type Result, ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import type { RawGraphMessage } from '../types/api-responses.js';
import type { EmailDetail } from '../types/teams.js';
import { requireGraphTokenAsync } from '../utils/auth-guards.js';
import { parseGraphEmail } from '../utils/parsers.js';
import { GRAPH_API } from '../utils/api-config.js';
import { graphAuthError } from './graph-files-api.js';

const MESSAGE_SELECT =
  'id,subject,from,sender,toRecipients,ccRecipients,bccRecipients,sentDateTime,receivedDateTime,' +
  'importance,isRead,webLink,conversationId,internetMessageId,body';
const ATTACHMENTS_EXPAND = 'attachments($select=name,contentType,size,isInline)';

/** Fetches one email by an `id` from teams_search_email results. Read-only. */
export async function getEmail(id: string): Promise<Result<EmailDetail>> {
  if (!id || /\s/.test(id)) {
    return err(createError(ErrorCode.INVALID_INPUT,
      'id must be a non-empty message id with no whitespace',
      { suggestions: ["Use an 'id' value from teams_search_email results"] }));
  }

  const tokenResult = await requireGraphTokenAsync();
  if (!tokenResult.ok) return tokenResult;
  const token = tokenResult.value;

  const params = new URLSearchParams({ $select: MESSAGE_SELECT, $expand: ATTACHMENTS_EXPAND });
  const response = await httpRequest<RawGraphMessage>(`${GRAPH_API.message(id)}?${params.toString()}`, {
    method: 'GET',
    headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
  });
  if (response.ok) return ok(parseGraphEmail(response.value.data));

  const authError = graphAuthError(response.error, token, 'email');
  if (authError) return err(authError);
  if (response.error.code !== ErrorCode.NOT_FOUND) return response;
  return err(createError(ErrorCode.NOT_FOUND,
    `Email not found. It may have been deleted, or the id belongs to a different mailbox. (${response.error.message})`,
    { retryable: false, suggestions: ["Use an 'id' value from teams_search_email results"] }));
}
