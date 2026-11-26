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
import { awaitRefreshQueue, refreshCoreTokensViaHttp } from '../auth/token-refresh.js';
import { createBrowserContext, closeBrowser } from '../browser/context.js';
import { ErrorCode, createError } from '../types/errors.js';
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

/**
 * Verifies messaging auth actually works after a sign-in before success is
 * reported. Sign-in alone does not mint skypetoken_asm — that cookie comes
 * from the HTTP Skype-token exchange — so when the session is not ready this
 * runs the same serialized refresh teams_login's fast path uses, then
 * requires messagingUsable(). Returns a failure ToolResult when messaging
 * stays unusable: reporting success here would loop login -> retry forever.
 */
async function messagingAuthFailure(): Promise<ToolResult | null> {
  let refreshError: string | undefined;
  if (!coreSessionReady()) {
    const refresh = await refreshCoreTokensViaHttp();
    if (!refresh.ok) {
      refreshError = refresh.error.message;
      log.warn('login', `Could not refresh credentials over HTTP: ${refreshError}`);
    }
  }
  if (messagingUsable()) return null;

  const detail = refreshError ? ` Refresh error: ${refreshError}` : '';
  return {
    success: false,
    error: createError(ErrorCode.API_ERROR,
      `Teams sign-in succeeded, but messaging credentials (skypetoken_asm) could not be minted.${detail} Do not call teams_login again in a loop; even with the user's approval, forceNew will not fix this: it only deletes local login data, and a fresh sign-in hits the same failure.`,
      {
        retryable: false,
        suggestions: [
          'Tell the user Teams sign-in succeeded but messaging authorization failed, and show them this error',
          'Do not call teams_login again in a loop — even with the user\'s approval, forceNew will not fix this failure',
          'Retry later, or check the network connection and Conditional Access policies',
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
    // Start from nothing: no profile, cookies, MSAL cache or session from any
    // earlier account can survive to be mixed into the new login.
    clearAllLocalData();
    clearRegionCache();
    return loginInVisibleBrowser();
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
  // and launching a second browser against it fails outright.
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

    // Sign-in succeeded but that does not prove messaging auth works: a
    // visible browser cannot mint skypetoken_asm either, so do NOT fall
    // through — report the failure instead of a false success.
    const failure = await messagingAuthFailure();
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

  return loginInVisibleBrowser();
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

  const failure = await messagingAuthFailure();
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
