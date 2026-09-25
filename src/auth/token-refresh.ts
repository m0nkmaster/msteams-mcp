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
 */

import { ErrorCode, createError } from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';
import {
  extractSubstrateToken,
  getValidAssignmentsToken,
  getValidGraphToken,
} from './token-extractor.js';
import { clearTokenCache } from './session-store.js';
import { refreshTokensViaHttp, type OnDemandResource } from './token-refresh-http.js';
import { createBrowserContext, closeBrowser } from '../browser/context.js';
import { ensureAuthenticated } from '../browser/auth.js';
import * as log from '../utils/logger.js';

/** Serialize refreshes because both HTTP and browser flows replace session state. */
let refreshQueue: Promise<unknown> = Promise.resolve();
let coreRefresh: Promise<Result<void>> | undefined;

function serializeRefresh<T>(refresh: () => Promise<T>): Promise<T> {
  const pending = refreshQueue.then(refresh, refresh);
  refreshQueue = pending.then(() => undefined, () => undefined);
  return pending;
}

/** Refresh core Teams credentials, with browser fallback. */
export function refreshTokensViaBrowser(): Promise<Result<void>> {
  if (!coreRefresh) {
    coreRefresh = serializeRefresh(refreshCoreTokens).finally(() => { coreRefresh = undefined; });
  }
  return coreRefresh;
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

async function refreshCoreTokens(): Promise<Result<void>> {
  // ── Strategy 1: HTTP refresh (fast, no browser needed) ──────────────
  // An expired access token does not block this: the MSAL refresh token in
  // session state is sufficient to acquire new resource tokens.
  const httpResult = await refreshTokensViaHttp();
  if (httpResult.ok && extractSubstrateToken()) return ok(undefined);

  log.warn('token-refresh', httpResult.ok
    ? 'HTTP refresh reported success but no valid Substrate token found, falling back to browser'
    : `HTTP refresh failed: ${httpResult.error.message}, falling back to browser (session cookies may still be valid)`);

  // ── Strategy 2: Browser refresh (fallback) ──────────────────────────
  return refreshViaHeadlessBrowser();
}

/**
 * Opens a headless browser with the persistent profile and lets MSAL
 * silently refresh tokens using its session cookies.
 */
async function refreshViaHeadlessBrowser(): Promise<Result<void>> {
  let manager: Awaited<ReturnType<typeof createBrowserContext>> | null = null;

  try {
    manager = await createBrowserContext({ headless: true });

    // headless: true fails fast if user interaction is required
    await ensureAuthenticated(manager.page, manager.context,
      msg => log.debug('token-refresh', msg), false, true);

    // ensureAuthenticated already saved the session
    await closeBrowser(manager, false);
    manager = null;
    clearTokenCache();

    if (!extractSubstrateToken()) {
      return err(createError(
        ErrorCode.AUTH_EXPIRED,
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
  }
}
