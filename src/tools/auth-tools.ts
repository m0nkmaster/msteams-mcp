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
  extractSkypeSpacesToken,
} from '../auth/token-extractor.js';
import { createError, ErrorCode } from '../types/errors.js';
import { refreshTokensViaHttp } from '../auth/token-refresh-http.js';
import { createBrowserContext, closeBrowser } from '../browser/context.js';
import * as log from '../utils/logger.js';
import { resetAssignmentsAvailability, clearRegionCache } from '../utils/auth-guards.js';
import { LOGIN_TIMEOUT_MS } from '../constants.js';
import { ensureAuthenticated, StaleTeamsSession, TeamsLoginTimeout } from '../browser/auth.js';

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
        description: 'Delete all previous login data (browser profile, cookies, saved session, every previously signed-in account) and sign in from scratch in a visible browser. Use only when the user wants to switch accounts or explicitly requests a full reset (default: false)',
      },
    },
  },
};

const statusToolDefinition: Tool = {
  name: 'teams_status',
  description: 'Check credential availability for each core Teams capability. session.exists reports whether the saved session file exists.',
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

export function getAuthCapabilities() {
  const substrate = getSubstrateTokenStatus();
  const messaging = getMessageAuthStatus();
  const canMessage = messaging.hasToken && extractMessageAuth() !== null;
  return {
    directApi: { available: substrate.hasToken, expiresAt: substrate.expiresAt, minutesRemaining: substrate.minutesRemaining },
    messaging: { available: canMessage, expiresAt: messaging.expiresAt, minutesRemaining: messaging.minutesRemaining },
    favorites: { available: canMessage && extractCsaToken() !== null },
    calendar: { available: canMessage && extractSkypeSpacesToken() !== null },
  };
}

async function finishLogin(message: string, tokenStatus?: ReturnType<typeof getSubstrateTokenStatus>, topUp = true): Promise<ToolResult> {
  let capabilities = getAuthCapabilities();
  let refreshError: string | undefined;
  if (topUp && Object.values(capabilities).some(capability => !capability.available)) {
    // Teams acquires some tokens lazily. Even a successful refresh can fill only some gaps.
    const refresh = await refreshTokensViaHttp();
    if (!refresh.ok) {
      refreshError = refresh.error.message;
      log.warn('login', `Could not top up core credentials: ${refreshError}`);
    }
    capabilities = getAuthCapabilities();
  }

  const data = { capabilities, ...(tokenStatus && { tokenStatus }) };
  const missing = Object.entries(capabilities).filter(([, value]) => !value.available).map(([name]) => name);
  if (missing.length) {
    const detail = `Unavailable capabilities: ${missing.join(', ')}.${refreshError ? ` Refresh failed: ${refreshError}` : ''}`;
    if (missing.length === Object.keys(capabilities).length) {
      return { success: false, error: createError(ErrorCode.AUTH_REQUIRED, `Login yielded no usable core credentials. ${detail}`) };
    }
    return { success: true, data: { message: 'Some Teams capabilities are unavailable.', warning: detail, ...data } };
  }

  return { success: true, data: { message, ...data } };
}

async function handleLogin(
  input: z.infer<typeof LoginInputSchema>,
): Promise<ToolResult> {
  // A fresh login may have changed what the account can access (e.g. new consent)
  resetAssignmentsAvailability();

  const deadline = Date.now() + LOGIN_TIMEOUT_MS;
  if (input.forceNew) {
    // Start from nothing: no profile, cookies, MSAL cache or session from any
    // earlier account can survive to be mixed into the new login.
    clearAllLocalData();
    clearRegionCache();
    return loginInVisibleBrowser(deadline);
  }

  // Fast path: if tokens are still valid, skip browser entirely
  const tokenStatus = getSubstrateTokenStatus();
  if (tokenStatus.hasToken && tokenStatus.minutesRemaining! >= TOKEN_VALID_THRESHOLD_MINUTES) {
    return finishLogin(`Already authenticated. Token valid for ${tokenStatus.minutesRemaining} more minutes.`, tokenStatus);
  }

  if (hasSessionState()) {
    const refreshed = await refreshTokensViaHttp();
    if (!refreshed.ok) log.info('login', `Saved-session refresh failed (${refreshed.error.code}). Trying browser SSO.`);
    if (refreshed.ok && getSubstrateTokenStatus().hasToken) {
      return finishLogin('Saved Teams session refreshed.', undefined, false);
    }
  }

  const headless = await createBrowserContext({ headless: true });
  try {
    await ensureAuthenticated(
      headless.page,
      headless.context,
      (msg) => log.info('login:headless', msg),
      false, // No overlay in headless
      true,
      deadline,
    );
  } catch (error) {
    log.warn('login:headless', `Headless SSO failed, falling back to visible browser: ${error instanceof Error ? error.message : String(error)}`);
    await closeBrowser(headless);
    return loginInVisibleBrowser(deadline, error instanceof StaleTeamsSession);
  }
  await closeBrowser(headless);
  return finishLogin('Login completed silently via SSO. Session has been saved.');
}

/** Opens a visible browser for the user to sign in, then saves the session. */
async function loginInVisibleBrowser(deadline: number, freshProfile = false): Promise<ToolResult> {
  const visible = await createBrowserContext({ headless: false, freshProfile });
  try {
    await ensureAuthenticated(visible.page, visible.context, (msg) => log.info('login', msg), true, false, deadline);
  } catch (error) {
    if (error instanceof TeamsLoginTimeout) {
      return { success: false, error: createError(ErrorCode.TIMEOUT, error.message) };
    }
    throw error;
  } finally {
    // Close browser after login - we only need the saved session/tokens
    await closeBrowser(visible);
  }

  return finishLogin('Login completed successfully. Session has been saved.');
}

async function handleStatus(): Promise<ToolResult> {
  return {
    success: true,
    data: {
      version: pkg.version,
      ...getAuthCapabilities(),
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
