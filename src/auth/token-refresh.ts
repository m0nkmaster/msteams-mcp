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

import { ErrorCode, createError, type McpError } from '../types/errors.js';
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
import { coreCredentialsComplete, coreSessionReady, fillCoreCredentialGaps } from './core-credentials.js';
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

/**
 * Number of teams_login browser sign-ins currently open on the persistent
 * profile. A second Chromium instance on the same profile fails on the
 * profile lock anyway, and a visible sign-in can run for minutes — so the
 * browser fallback stands down for the whole window instead of colliding.
 * Counted, not boolean: handleLogin is not serialized, so overlapping
 * logins must not clear the guard while another browser is still open.
 */
let browserLoginCount = 0;

/** Marks a teams_login browser session so refreshViaHeadlessBrowser does not collide on the profile. */
export function setBrowserLoginInProgress(active: boolean): void {
  browserLoginCount = Math.max(0, browserLoginCount + (active ? 1 : -1));
}

/** True while at least one teams_login browser sign-in is in progress. */
export function isBrowserLoginInProgress(): boolean {
  return browserLoginCount > 0;
}

/** Shared error for refreshes that stand down while a sign-in owns the profile. */
function loginInProgressError(): McpError {
  return createError(
    ErrorCode.AUTH_REQUIRED,
    'Teams sign-in is already in progress in another window; wait for it to finish, then retry.',
    { retryable: false, suggestions: ['Wait for the in-progress teams_login sign-in to complete, then retry the tool call'] },
  );
}

function serializeRefresh<T>(refresh: () => Promise<T>): Promise<T> {
  const run = () => withTimeout(refresh(), CORE_REFRESH_TIMEOUT_MS, 'token refresh');
  const pending = refreshQueue.then(run, run);
  refreshQueue = pending.then(() => undefined, () => undefined);
  return pending;
}

/**
 * Resolves once the current refresh queue has settled AND any in-flight
 * core refresh has finished its catch-path cleanup. teams_login awaits
 * this before opening a browser so it cannot collide with the Chromium
 * profile lock held by an in-flight refresh — or with the wedged-browser
 * kill that runs after the refresh deadline fires. Never rejects.
 */
export async function awaitRefreshQueue(): Promise<void> {
  // Loop until the queue is stable: a refresh enqueued while we waited
  // chains on the same tail and would otherwise slip past a one-shot
  // snapshot — its session-state write could land after a forceNew wipe.
  for (;;) {
    const tail = refreshQueue;
    const core = coreRefresh;
    await Promise.all([tail, core ?? Promise.resolve()]);
    if (refreshQueue === tail && coreRefresh === undefined) return;
  }
}

/** Refresh core Teams credentials, with browser fallback. */
export function refreshTokensViaBrowser(): Promise<Result<void>> {
  // Entry gate, not just the browser fallback: the HTTP leg writes session
  // state, and a write landing mid-sign-in (or mid-forceNew wipe) can
  // resurrect the old account's session.
  if (isBrowserLoginInProgress()) {
    return Promise.resolve(err(loginInProgressError()));
  }
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
 * Core-credential HTTP refresh serialized with every other refresh. Callers
 * outside the refresh queue (e.g. the teams_login fast path) must not race
 * session-state writes against an in-flight refreshTokensViaBrowser.
 */
export function refreshCoreTokensViaHttp(): Promise<Result<void>> {
  // Re-check inside the serialized job: a refresh queued behind an in-flight
  // one may already have left the session ready, making a second OAuth
  // exchange redundant. Extractor caches are invalidated by the refresh's
  // own writeSessionState + clearTokenCache, so this sees fresh state.
  return serializeRefresh(() => coreSessionReady()
    ? Promise.resolve(ok(undefined))
    : refreshTokensViaHttp()).catch((error): Result<void> => {
    if (error instanceof TimeoutError) {
      return err(createError(ErrorCode.TIMEOUT,
        `Core credential refresh timed out after ${CORE_REFRESH_TIMEOUT_MS / 1000}s.`));
    }
    const message = error instanceof Error ? error.message : String(error);
    return err(createError(ErrorCode.UNKNOWN,
      `Core credential refresh failed: ${message}`));
  });
}

/**
 * Optional resources whose last on-demand refresh failed with an auth-state
 * error (the AUTH_EXPIRED/AUTH_REQUIRED branch below). Those need a real
 * sign-in — headless SSO re-mints the SPA refresh token — so teams_login's
 * fast path must not report "Already authenticated" while one is flagged.
 * ACCESS_DENIED, TIMEOUT, API_ERROR and the in-progress gate never flag:
 * sign-in cannot fix consent, Conditional Access or a wedged network.
 * Per process: the long-lived MCP server sees the flag; a fresh CLI login
 * process starts empty and keeps the fast path.
 */
const optionalSignInNeeded = new Set<OnDemandResource>();

function optionalResourceConfig(resource: OnDemandResource): { label: string; getValid: () => string | null } {
  // Resolved at call time, not module scope: partial token-extractor mocks in
  // unrelated test files must not break importing this module.
  return resource === 'graph'
    ? { label: 'Microsoft Graph', getValid: getValidGraphToken }
    : { label: 'Assignments', getValid: getValidAssignmentsToken };
}

/**
 * Optional resources currently flagged as needing a sign-in. A flag clears
 * itself once the session holds an unexpired token again: the on-demand
 * guard serves that token without any exchange, so a stale flag must not
 * keep sending teams_login down the sign-in path.
 */
export function optionalResourcesNeedingSignIn(): OnDemandResource[] {
  for (const resource of [...optionalSignInNeeded]) {
    if (optionalResourceConfig(resource).getValid()) {
      optionalSignInNeeded.delete(resource);
    }
  }
  return [...optionalSignInNeeded];
}

/**
 * Acquire an optional resource's token on demand with a single HTTP exchange.
 * Optional features never launch a browser, refresh core credentials, or return
 * an error that would trigger the server's auto-login.
 */
function onDemandExchange(resource: OnDemandResource, label: string, getValid: () => string | null, viaLogin = false): Promise<Result<string>> {
  return serializeRefresh(async () => {
    if (viaLogin) {
      // Sign-in just ran: a token the session already holds needs no
      // exchange, and whatever the outcome the flag is spent — a later
      // on-demand auth-state failure re-arms it, so a resource sign-in
      // cannot fix stops forcing a sign-in on every login.
      const cached = getValid();
      if (cached) {
        optionalSignInNeeded.delete(resource);
        return ok(cached);
      }
    }
    const result = await refreshTokensViaHttp(resource);
    if (viaLogin) optionalSignInNeeded.delete(resource);
    if (!result.ok && (result.error.code === ErrorCode.AUTH_EXPIRED || result.error.code === ErrorCode.AUTH_REQUIRED)) {
      if (!viaLogin) optionalSignInNeeded.add(resource);
      return err(createError(ErrorCode.AUTH_INTERACTION_REQUIRED,
        `${label} could not be authorized: ${result.error.message}`,
        { retryable: false }));
    }
    if (!result.ok) return result;
    optionalSignInNeeded.delete(resource);
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

function refreshOnDemandToken(resource: OnDemandResource, label: string, getValid: () => string | null): Promise<Result<string>> {
  // Same entry gate as the core refresh: the exchange writes session state.
  // Optional resources never return AUTH_REQUIRED/AUTH_EXPIRED (that would
  // trigger the server's auto-login), so this maps to AUTH_INTERACTION_REQUIRED.
  if (isBrowserLoginInProgress()) {
    return Promise.resolve(err(createError(ErrorCode.AUTH_INTERACTION_REQUIRED,
      `Teams sign-in is already in progress in another window; the ${label} token is unavailable until it finishes.`,
      { retryable: false })));
  }
  return onDemandExchange(resource, label, getValid);
}

/**
 * Re-mint an optional token right after a sign-in: teams_login owns the
 * profile for the whole window, so the isBrowserLoginInProgress entry gate
 * must not stand this down (same reasoning as refreshCoreTokensViaHttp).
 */
export function refreshOptionalTokenForLogin(resource: OnDemandResource): Promise<Result<string>> {
  const cfg = optionalResourceConfig(resource);
  return onDemandExchange(resource, cfg.label, cfg.getValid, true);
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
export function coreSessionUsable(): boolean {
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
  // teams_login owns the profile for the whole sign-in window: a second
  // browser on it fails on the Chromium lock, so fail fast instead.
  if (isBrowserLoginInProgress()) {
    return err(loginInProgressError());
  }

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
