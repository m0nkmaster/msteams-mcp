/**
 * Authentication-related tool handlers.
 */

import { createRequire } from 'module';
import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolResult } from './index.js';

const require = createRequire(import.meta.url);
const pkg = require('../../package.json') as { version: string };
import {
  hasSessionState,
  isSessionLikelyExpired,
  clearAllLocalData,
} from '../auth/session-store.js';
import {
  getSubstrateTokenStatus,
  getMessageAuthStatus,
  extractMessageAuth,
  extractCsaToken,
} from '../auth/token-extractor.js';
import { coreSessionReady, messagingUsable } from '../auth/core-credentials.js';
import { awaitRefreshQueue, refreshCoreTokensViaHttp, coreSessionUsable, setBrowserLoginInProgress, isBrowserLoginInProgress } from '../auth/token-refresh.js';
import { createBrowserContext, closeBrowser } from '../browser/context.js';
import { ErrorCode, createError, type McpError } from '../types/errors.js';
import { TOKEN_VALID_THRESHOLD_MINUTES } from '../constants.js';
import * as log from '../utils/logger.js';
import { resetAssignmentsAvailability, clearRegionCache } from '../utils/auth-guards.js';
import { ensureAuthenticated } from '../browser/auth.js';

// ─────────────────────────────────────────────────────────────────────────────
// Schemas
// ─────────────────────────────────────────────────────────────────────────────

export const LoginInputSchema = z.object({
  forceNew: z.boolean().optional().default(false),
});

// ─────────────────────────────────────────────────────────────────────────────
// Tool Definitions
// ─────────────────────────────────────────────────────────────────────────────

const loginToolDefinition: Tool = {
  name: 'teams_login',
  description: 'Sign in to Microsoft Teams or refresh an expired session. Call without forceNew for normal re-login: it tries silent SSO first and opens a browser window only if needed.',
  inputSchema: {
    type: 'object',
    properties: {
      forceNew: {
        type: 'boolean',
        description: 'DESTRUCTIVE: deletes all local login data (browser profile, cookies, saved session, every previously signed-in account) and forces an interactive sign-in in a visible browser. Use only with the user\'s approval, after a plain teams_login has failed or signed in the wrong account, or to switch accounts (default: false)',
      },
    },
  },
};

const statusToolDefinition: Tool = {
  name: 'teams_status',
  description: 'Check the current authentication status and session state.',
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// Handlers
// ─────────────────────────────────────────────────────────────────────────────

/** Refresh failures that mean the saved sign-in state itself is broken. */
const AUTH_STATE_CODES: ReadonlySet<ErrorCode> = new Set([
  ErrorCode.AUTH_REQUIRED,
  ErrorCode.AUTH_EXPIRED,
  ErrorCode.AUTH_INTERACTION_REQUIRED,
]);

/**
 * Verifies the session is actually usable after a sign-in before success is
 * reported. Sign-in alone does not mint skypetoken_asm — that cookie comes
 * from the HTTP Skype-token exchange — and refreshTokensViaHttp tolerates a
 * failed scope, so messaging can work while Substrate is missing. When the
 * session is not ready this runs the same serialized refresh teams_login's
 * fast path uses, then requires coreSessionUsable() AND messagingUsable().
 * Returns a failure ToolResult otherwise: reporting success here would loop
 * login -> retry forever.
 */
async function messagingAuthFailure(forceNewMayHelp: boolean): Promise<ToolResult | null> {
  let refreshError: McpError | undefined;
  if (!coreSessionReady()) {
    const refresh = await refreshCoreTokensViaHttp();
    if (!refresh.ok) {
      refreshError = refresh.error;
      log.warn('login', `Could not refresh credentials over HTTP: ${refreshError.message}`);
    }
  }
  if (coreSessionUsable() && messagingUsable()) return null;

  const detail = refreshError
    ? ` Refresh error: ${refreshError.message}`
    : ' The Skype token exchange failed or returned no cookie; see the server log for details.';

  if (!messagingUsable()) {
    // AUTH_* refresh failures mean the local auth state itself is broken.
    // forceNew (a fresh sign-in) can only help when the sign-in that just
    // ran was headless SSO reusing old state — after an interactive
    // sign-in it would repeat the same wipe and cannot fix anything.
    if (forceNewMayHelp && refreshError !== undefined && AUTH_STATE_CODES.has(refreshError.code)) {
      return {
        success: false,
        error: createError(ErrorCode.API_ERROR,
          `Teams sign-in succeeded, but messaging credentials (skypetoken_asm) could not be minted.${detail} Do not call teams_login again in a loop; only with the user's approval, teams_login with { "forceNew": true } may help: it deletes all local login data and starts a fresh sign-in.`,
          {
            retryable: false,
            suggestions: [
              'Tell the user Teams sign-in succeeded but messaging authorization failed, and show them this error',
              'Only with the user\'s approval, call teams_login with { "forceNew": true } to clear the saved sign-in state and retry',
              'Retry later after checking the network connection and Conditional Access policies; teams_login does not retry on its own',
            ],
          }),
      };
    }
    return {
      success: false,
      error: createError(ErrorCode.API_ERROR,
        `Teams sign-in succeeded, but messaging credentials (skypetoken_asm) could not be minted.${detail} Do not call teams_login again in a loop; even with the user's approval, forceNew will not fix this: it only deletes local login data, and a fresh sign-in hits the same failure.`,
        {
          retryable: false,
          suggestions: [
            'Tell the user Teams sign-in succeeded but messaging authorization failed, and show them this error',
            'Do not call teams_login again in a loop; even with the user\'s approval, forceNew will not fix this failure',
            'Retry later after checking the network connection and Conditional Access policies; teams_login does not retry on its own',
          ],
        }),
    };
  }

  // Messaging works but Substrate is missing: the next tool call would fail
  // auth and bounce straight back to login, so report the gap instead. The
  // detail names the actual gap — the Skype exchange is not the cause here.
  const substrateDetail = refreshError
    ? ` Refresh error: ${refreshError.message}`
    : ' The Substrate token exchange failed or returned no token; see the server log for details.';
  return {
    success: false,
    error: createError(ErrorCode.API_ERROR,
      `Teams sign-in succeeded, but the Substrate session token is missing.${substrateDetail} Do not call teams_login again in a loop; a later refresh may mint it.`,
      {
        retryable: false,
        suggestions: [
          'Tell the user Teams sign-in succeeded but the session is incomplete, and show them this error',
          'Do not call teams_login again in a loop; retry the failed tool call later instead',
          'If it keeps failing, check the network connection and Conditional Access policies',
        ],
      }),
  };
}

/** Failure returned when a teams_login browser sign-in is already running. */
function loginInProgressFailure(): ToolResult {
  return {
    success: false,
    error: createError(ErrorCode.API_ERROR,
      'A Teams sign-in is already in progress in another window. Wait for it to finish, then retry teams_login.',
      {
        retryable: false,
        suggestions: [
          'Wait for the in-progress teams_login sign-in window to complete, then retry',
          'If no sign-in window is visible, the earlier attempt may have stalled; call teams_login again once it clears',
        ],
      }),
  };
}

async function handleLogin(
  input: z.infer<typeof LoginInputSchema>,
): Promise<ToolResult> {
  // A fresh login may have changed what the account can access (e.g. new consent)
  resetAssignmentsAvailability();

  if (input.forceNew) {
    // A second sign-in would collide with the open one on the profile lock,
    // and its cleanup could drop the guard under the first browser.
    if (isBrowserLoginInProgress()) return loginInProgressFailure();
    // The guard rises before anything touches the profile: an auto-refresh
    // starting while the queue drains or the directory is deleted must not
    // launch a browser onto it.
    setBrowserLoginInProgress(true);
    try {
      // Wait out any in-flight refresh first: it may be holding the browser
      // profile that clearAllLocalData is about to delete.
      await awaitRefreshQueue();
      // Start from nothing: no profile, cookies, MSAL cache or session from any
      // earlier account can survive to be mixed into the new login.
      clearAllLocalData();
      clearRegionCache();
      return await loginInVisibleBrowser();
    } finally {
      setBrowserLoginInProgress(false);
    }
  }

  // Fast path: if tokens are still valid, skip browser entirely — but only
  // when messaging auth is usable too. One HTTP refresh runs first: it can
  // re-mint an expired skypetoken_asm without opening a browser.
  const tokenStatus = getSubstrateTokenStatus();
  if (tokenStatus.hasToken && tokenStatus.minutesRemaining! >= TOKEN_VALID_THRESHOLD_MINUTES) {
    if (!coreSessionReady()) {
      const refresh = await refreshCoreTokensViaHttp();
      if (!refresh.ok) {
        log.warn('login', `Could not refresh credentials over HTTP: ${refresh.error.message}`);
      }
    }
    if (messagingUsable()) {
      const refreshedStatus = getSubstrateTokenStatus();
      return {
        success: true,
        data: {
          message: `Already authenticated. Token valid for ${refreshedStatus.minutesRemaining} more minutes.`,
          tokenStatus: {
            expiresAt: refreshedStatus.expiresAt,
            minutesRemaining: refreshedStatus.minutesRemaining,
          },
        },
      };
    }
    log.warn('login', 'Substrate token is valid but messaging auth is unusable; continuing to browser login');
  }

  // Headless-first strategy:
  // The persistent browser profile retains Microsoft's long-lived session cookies,
  // so headless SSO can succeed even without a session-state file.
  // Wait out any in-flight refresh first: it holds the Chromium profile lock,
  // and launching a second browser against it fails outright. The flag stays
  // up for the whole sign-in (headless, fallback gap, visible browser) so an
  // auto-refresh cannot slip a browser onto the profile mid-login. A second
  // login call stands down instead of colliding on the profile lock.
  if (isBrowserLoginInProgress()) return loginInProgressFailure();
  setBrowserLoginInProgress(true);
  try {
    await awaitRefreshQueue();
    const headless = await createBrowserContext({ headless: true });
    try {
      await ensureAuthenticated(
        headless.page,
        headless.context,
        (msg) => log.info('login:headless', msg),
        false, // No overlay in headless
        true   // Headless mode - throw immediately if user interaction required
      );
      await closeBrowser(headless, true);

      // Sign-in succeeded but that does not prove the session is usable: a
      // visible browser cannot mint skypetoken_asm either, so do NOT fall
      // through — report the failure instead of a false success. Headless
      // SSO reused the old profile state, so an auth-state refresh failure
      // here can still be fixed by forceNew; after the visible sign-in below
      // it cannot (messagingAuthFailure(false)).
      const failure = await messagingAuthFailure(true);
      if (failure) return failure;

      return {
        success: true,
        data: {
          message: 'Login completed silently via SSO. Session has been saved.',
        },
      };
    } catch (error) {
      log.warn('login:headless', `Headless SSO failed, falling back to visible browser: ${error instanceof Error ? error.message : String(error)}`);
      await closeBrowser(headless, false).catch(() => {});
    }

    return await loginInVisibleBrowser();
  } finally {
    setBrowserLoginInProgress(false);
  }
}

/** Opens a visible browser for the user to sign in, then saves the session. */
async function loginInVisibleBrowser(): Promise<ToolResult> {
  const visible = await createBrowserContext({ headless: false });
  try {
    await ensureAuthenticated(visible.page, visible.context, (msg) => log.info('login', msg));
  } finally {
    // Close browser after login - we only need the saved session/tokens
    await closeBrowser(visible, true);
  }

  const failure = await messagingAuthFailure(false);
  if (failure) return failure;

  return {
    success: true,
    data: {
      message: 'Login completed successfully. Session has been saved.',
    },
  };
}

async function handleStatus(): Promise<ToolResult> {
  const tokenStatus = getSubstrateTokenStatus();
  const messageAuthStatus = getMessageAuthStatus();

  return {
    success: true,
    data: {
      version: pkg.version,
      directApi: {
        available: tokenStatus.hasToken,
        expiresAt: tokenStatus.expiresAt,
        minutesRemaining: tokenStatus.minutesRemaining,
      },
      messaging: {
        available: messageAuthStatus.hasToken,
        expiresAt: messageAuthStatus.expiresAt,
        minutesRemaining: messageAuthStatus.minutesRemaining,
      },
      favorites: {
        available: extractMessageAuth() !== null && extractCsaToken() !== null,
      },
      session: {
        exists: hasSessionState(),
        likelyExpired: isSessionLikelyExpired(),
      },
    },
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Exports
// ─────────────────────────────────────────────────────────────────────────────

export const loginTool: RegisteredTool<typeof LoginInputSchema> = {
  definition: loginToolDefinition,
  schema: LoginInputSchema,
  handler: handleLogin,
};

export const statusTool: RegisteredTool<z.ZodObject<Record<string, never>>> = {
  definition: statusToolDefinition,
  schema: z.object({}),
  handler: handleStatus,
};

/** All auth-related tools. */
export const authTools = [loginTool, statusTool];
