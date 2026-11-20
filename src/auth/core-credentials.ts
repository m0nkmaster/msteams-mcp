/**
 * Core credential completeness check and HTTP top-up.
 *
 * Shared by teams_login and the automatic token-refresh path: a session is
 * only usable when every core credential is present, and the browser can
 * leave some of them unminted.
 */

import {
  extractSubstrateToken,
  extractSkypeSpacesToken,
  extractCsaToken,
  extractMessageAuth,
} from './token-extractor.js';
import { refreshTokensViaHttp } from './token-refresh-http.js';
import * as log from '../utils/logger.js';

/** Every core credential a session needs: search, calendar, favourites, messaging. */
export function coreCredentialsComplete(): boolean {
  return extractSubstrateToken() !== null
    && extractSkypeSpacesToken() !== null
    && extractCsaToken() !== null
    && extractMessageAuth() !== null;
}

/**
 * Mints any core credential the browser did not leave in the saved session.
 *
 * Teams JS acquires some resource tokens lazily — the CSA (chatsvcagg) token
 * behind favourites is routinely absent from a session saved straight after
 * SSO, so login would report success while teams_status reported favourites
 * unavailable until some later call happened to trigger a refresh. One HTTP
 * exchange mints every core resource. Failure is non-fatal: the session is
 * still usable and the next tool call retries through auto-login.
 */
export async function fillCoreCredentialGaps(logContext: string = 'login'): Promise<void> {
  if (coreCredentialsComplete()) return;

  const result = await refreshTokensViaHttp();
  if (!result.ok) {
    log.warn(logContext, `Could not top up core credentials: ${result.error.message}`);
  }
}
