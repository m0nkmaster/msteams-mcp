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
  extractSubstrateToken,
  extractSkypeSpacesToken,
} from '../auth/token-extractor.js';
import { refreshTokensViaHttp } from '../auth/token-refresh-http.js';
import { createBrowserContext, closeBrowser } from '../browser/context.js';
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
  description: 'Trigger manual login flow for Microsoft Teams. Use this if the session has expired or you need to switch accounts.',
  inputSchema: {
    type: 'object',
    properties: {
      forceNew: {
        type: 'boolean',
        description: 'Delete all previous login data (browser profile, cookies, saved session, every previously signed-in account) and sign in from scratch in a visible browser. Use to switch accounts or recover a broken session (default: false)',
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

/** Minimum minutes remaining on token to consider it valid (skip browser). */
const TOKEN_VALID_THRESHOLD_MINUTES = 10;

/** Every core credential a session needs: search, calendar, favourites, messaging. */
function coreCredentialsComplete(): boolean {
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
async function fillCoreCredentialGaps(): Promise<void> {
  if (coreCredentialsComplete()) return;

  const result = await refreshTokensViaHttp();
  if (!result.ok) {
    log.warn('login', `Could not top up core credentials: ${result.error.message}`);
  }
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

  // Fast path: if tokens are still valid, skip browser entirely
  const tokenStatus = getSubstrateTokenStatus();
  if (tokenStatus.hasToken && tokenStatus.minutesRemaining! >= TOKEN_VALID_THRESHOLD_MINUTES) {
    await fillCoreCredentialGaps();
    return {
      success: true,
      data: {
        message: `Already authenticated. Token valid for ${tokenStatus.minutesRemaining} more minutes.`,
        tokenStatus: {
          expiresAt: tokenStatus.expiresAt,
          minutesRemaining: tokenStatus.minutesRemaining,
        },
      },
    };
  }

  // Headless-first strategy:
  // The persistent browser profile retains Microsoft's long-lived session cookies,
  // so headless SSO can succeed even without a session-state file.
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
    await fillCoreCredentialGaps();

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

  await fillCoreCredentialGaps();

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
