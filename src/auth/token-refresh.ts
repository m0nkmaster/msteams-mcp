/**
 * Token refresh orchestrator.
 * 
 * Tries HTTP-based refresh first (direct OAuth2 token endpoint call, ~100ms),
 * then falls back to headless browser refresh (~8s) if HTTP fails.
 * 
 * HTTP refresh works by extracting the MSAL refresh token from session state
 * and exchanging it for new access tokens via Azure AD's token endpoint.
 * This works identically for standard Microsoft login and corporate SSO
 * (ADFS/Okta federation) — refresh tokens are standard Azure AD tokens
 * regardless of how the user originally authenticated.
 * 
 * Browser fallback covers cases where:
 * - The refresh token has expired (typically after days/weeks of inactivity)
 * - Conditional Access policies require interactive auth
 * - The MSAL cache format has changed unexpectedly
 * 
 * First login always requires a browser — there's no refresh token to use yet.
 * 
 * Every serialized refresh is bounded by CORE_REFRESH_TIMEOUT_MS: a wedged
 * browser or network call rejects its waiters instead of hanging the server,
 * the queue still advances, and the next call starts a fresh attempt.
 */

import { ErrorCode, createError } from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';
import {
  extractSubstrateToken,
  extractMessageAuth,
  getValidAssignmentsToken,
  getValidGraphToken,
} from './token-extractor.js';
import { clearTokenCache } from './session-store.js';
import { refreshTokensViaHttp, type OnDemandResource } from './token-refresh-http.js';
import { createBrowserContext, closeBrowser, type BrowserManager } from '../browser/context.js';
import { ensureAuthenticated } from '../browser/auth.js';
import { coreCredentialsComplete, fillCoreCredentialGaps } from './core-credentials.js';
import { withTimeout, TimeoutError } from '../utils/timeout.js';
import { CORE_REFRESH_TIMEOUT_MS, BROWSER_REFRESH_LAUNCH_TIMEOUT_MS } from '../constants.js';
import * as log from '../utils/logger.js';

/** Serialize refreshes because both HTTP and browser flows replace session state. */
let refreshQueue: Promise<unknown> = Promise.resolve();
let coreRefresh: Promise<Result<void>> | undefined;

/**
 * The headless browser launched by the in-flight core refresh, if any.
 * Tracked so the refresh deadline can release the profile lock it holds.
 */
let activeBrowser: BrowserManager | null = null;

function serializeRefresh<T>(refresh: () => Promise<T>): Promise<T> {
  const run = () => withTimeout(refresh(), CORE_REFRESH_TIMEOUT_MS, 'token refresh');
  const pending = refreshQueue.then(run, run);
  refreshQueue = pending.then(() => undefined, () => undefined);
  return pending;
}

/** Refresh core Teams credentials, with browser fallback. */
export function refreshTokensViaBrowser(): Promise<Result<void>> {
  if (!coreRefresh) {
    coreRefresh = serializeRefresh(refreshCoreTokens)
      .catch(async (error): Promise<Result<void>> => {
        if (error instanceof TimeoutError) {
          log.error('token-refresh', `Automatic token refresh timed out after ${CORE_REFRESH_TIMEOUT_MS / 1000}s; abandoning the attempt`);
          await cleanupWedgedBrowser();
          return err(createError(
            ErrorCode.AUTH_REQUIRED,
            `Automatic token refresh timed out after ${CORE_REFRESH_TIMEOUT_MS / 1000}s. Call teams_login to re-authenticate.`,
          ));
        }
        const message = error instanceof Error ? error.message : String(error);
        return err(createError(
          ErrorCode.UNKNOWN,
          `Token refresh failed: ${message}. Call teams_login to re-authenticate.`,
        ));
      })
      .finally(() => { coreRefresh = undefined; });
  }
  return coreRefresh;
}

/**
 * Best-effort cleanup after the refresh deadline fires: the abandoned attempt
 * may still hold the persistent browser profile open, which would make every
 * later refresh fail on the Chromium profile lock. closeBrowser bounds the
 * close itself and force-kills profile processes if it wedges.
 */
async function cleanupWedgedBrowser(): Promise<void> {
  const manager = activeBrowser;
  activeBrowser = null;
  if (manager) {
    await closeBrowser(manager, false).catch(() => {});
  }
}

/**
 * Acquire an optional resource's token on demand with a single HTTP exchange.
 * Optional features never launch a browser, refresh core credentials, or return
 * an error that would trigger the server's auto-login.
 */
function refreshOnDemandToken(resource: OnDemandResource, label: string, getValid: () => string | null): Promise<Result<string>> {
  return serializeRefresh(async () => {
    const result = await refreshTokensViaHttp(resource);
    if (!result.ok && (result.error.code === ErrorCode.AUTH_EXPIRED || result.error.code === ErrorCode.AUTH_REQUIRED)) {
      return err(createError(ErrorCode.AUTH_INTERACTION_REQUIRED,
        `${label} could not be authorized: ${result.error.message}`,
        { retryable: false }));
    }
    if (!result.ok) return result;
    const token = getValid();
    return token ? ok(token) : err(createError(ErrorCode.API_ERROR,
      `${label} token exchange did not return a valid token for the ${label} service.`,
      { retryable: false }));
  }).catch((error): Result<string> => {
    if (error instanceof TimeoutError) {
      return err(createError(ErrorCode.TIMEOUT,
        `${label} token refresh timed out after ${CORE_REFRESH_TIMEOUT_MS / 1000}s.`));
    }
    const message = error instanceof Error ? error.message : String(error);
    return err(createError(ErrorCode.UNKNOWN,
      `${label} token refresh failed: ${message}`));
  });
}

/** Acquire the optional EDU Assignments token. */
export function refreshAssignmentsToken(): Promise<Result<string>> {
  return refreshOnDemandToken('assignments', 'Assignments', getValidAssignmentsToken);
}

/** Acquire the optional Microsoft Graph token (used for file downloads). */
export function refreshGraphToken(): Promise<Result<string>> {
  return refreshOnDemandToken('graph', 'Microsoft Graph', getValidGraphToken);
}

/**
 * True when the credentials a retried tool call actually needs are usable:
 * a live Substrate token and the messaging cookies. Stricter gaps (e.g. a
 * missing CSA token) are repaired by the top-up but do not fail the refresh.
 */
function coreSessionUsable(): boolean {
  return extractSubstrateToken() !== null && extractMessageAuth() !== null;
}

async function refreshCoreTokens(): Promise<Result<void>> {
  // ── Strategy 1: HTTP refresh (fast, no browser needed) ──────────────
  // An expired access token does not block this: the MSAL refresh token in
  // session state is sufficient to acquire new resource tokens.
  const httpResult = await refreshTokensViaHttp();
  if (httpResult.ok) {
    if (coreCredentialsComplete()) return ok(undefined);

    // The refresh reported success but the session is incomplete — the Skype
    // cookie exchange is best-effort inside it, so messaging cookies can be
    // missing. Run the same one-shot top-up teams_login uses, then re-check.
    await fillCoreCredentialGaps('token-refresh');
    if (coreSessionUsable()) return ok(undefined);

    log.warn('token-refresh', 'HTTP refresh left core credentials incomplete, falling back to browser');
  } else {
    log.warn('token-refresh', `HTTP refresh failed: ${httpResult.error.message}, falling back to browser (session cookies may still be valid)`);
  }

  // ── Strategy 2: Browser refresh (fallback) ──────────────────────────
  return refreshViaHeadlessBrowser();
}

/**
 * Opens a headless browser with the persistent profile and lets MSAL
 * silently refresh tokens using its session cookies.
 */
async function refreshViaHeadlessBrowser(): Promise<Result<void>> {
  let manager: BrowserManager | null = null;

  try {
    manager = await createBrowserContext({ headless: true, launchTimeoutMs: BROWSER_REFRESH_LAUNCH_TIMEOUT_MS });
    activeBrowser = manager;

    // headless: true fails fast if user interaction is required
    await ensureAuthenticated(manager.page, manager.context,
      msg => log.debug('token-refresh', msg), false, true);

    // ensureAuthenticated already saved the session
    await closeBrowser(manager, false);
    manager = null;
    activeBrowser = null;
    clearTokenCache();

    // A browser save only contains what Teams JS had cached: the messaging
    // cookies are minted by the HTTP Skype-token exchange, so "successful"
    // SSO can still leave the session unusable. Top up once, exactly like
    // teams_login does, then require the credentials the retried call needs.
    if (!coreCredentialsComplete()) {
      log.info('token-refresh', 'Browser refresh left core credentials incomplete, running HTTP top-up');
      await fillCoreCredentialGaps('token-refresh');
    }

    if (!coreSessionUsable()) {
      return err(createError(
        ErrorCode.AUTH_REQUIRED,
        'ACTION REQUIRED: Token refresh failed. You MUST call teams_login to re-authenticate.',
      ));
    }
    return ok(undefined);
  } catch (error) {
    if (manager) {
      await closeBrowser(manager, false).catch(() => {});
    }
    const message = error instanceof Error ? error.message : 'Unknown error';
    return err(createError(
      ErrorCode.UNKNOWN,
      `Token refresh via browser failed: ${message}. Call teams_login to re-authenticate.`,
    ));
  } finally {
    if (manager && activeBrowser === manager) {
      activeBrowser = null;
    }
  }
}
