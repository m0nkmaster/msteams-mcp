/**
 * Token refresh via headless browser.
 * 
 * Teams uses SPA OAuth2 which restricts refresh tokens to browser-based CORS
 * requests. We open a headless browser with the persistent profile, let MSAL
 * silently refresh tokens using the profile's long-lived session cookies,
 * then save the updated state. Seamless to the user — no browser window shown.
 * 
 * The persistent profile is shared with visible login, so only one browser
 * can use it at a time (Chromium profile lock). A module-level flag prevents
 * concurrent refresh attempts.
 */

import { TOKEN_REFRESH_THRESHOLD_MS } from '../constants.js';
import { ErrorCode, createError } from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';
import {
  extractSubstrateToken,
  clearTokenCache,
} from './token-extractor.js';

/** Result of a successful token refresh. */
export interface TokenRefreshResult {
  /** New token expiry time. */
  newExpiry: Date;
  /** Previous expiry time (for comparison). */
  previousExpiry: Date;
  /** Minutes gained by refresh. */
  minutesGained: number;
  /** Whether a refresh was actually needed (token was close to expiry). */
  refreshNeeded: boolean;
}

/** Module-level flag to prevent concurrent browser profile access. */
let refreshInProgress = false;

/**
 * Refreshes tokens by opening a headless browser with the persistent profile.
 * The profile's long-lived Microsoft session cookies enable silent re-auth
 * without user interaction. MSAL only refreshes tokens when an API call
 * requires them, so we trigger a search via ensureAuthenticated.
 */
export async function refreshTokensViaBrowser(): Promise<Result<TokenRefreshResult>> {
  // Prevent concurrent access to the shared browser profile
  if (refreshInProgress) {
    return err(createError(
      ErrorCode.UNKNOWN,
      'Token refresh already in progress. Please wait and try again.',
      { retryable: true, suggestions: ['Wait a moment and retry the request'] }
    ));
  }

  // Get current token expiry for comparison
  const beforeToken = extractSubstrateToken();
  if (!beforeToken) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'No token found in session. Please run teams_login to authenticate.',
      { suggestions: ['Call teams_login to authenticate'] }
    ));
  }

  const previousExpiry = beforeToken.expiry;

  // Import browser functions dynamically to avoid circular dependencies
  const { createBrowserContext, closeBrowser } = await import('../browser/context.js');

  let manager: Awaited<ReturnType<typeof createBrowserContext>> | null = null;
  refreshInProgress = true;

  try {
    // Open headless browser with persistent profile
    manager = await createBrowserContext({ headless: true });

    // Import auth functions
    const { ensureAuthenticated } = await import('../browser/auth.js');

    // Use the same auth flow that works for login - this triggers token acquisition
    // showOverlay: false since headless browser has no visible window
    // headless: true to fail fast if user interaction is required
    await ensureAuthenticated(manager.page, manager.context, (msg) => {
      // Silent logging for headless refresh
      console.log(`[token-refresh] ${msg}`);
    }, false, true);

    // Close browser (ensureAuthenticated already saved the session)
    await closeBrowser(manager, false);
    manager = null;

    // Clear our token cache to force re-extraction from the new session
    clearTokenCache();

    // Extract the new token to verify we still have valid tokens
    const afterToken = extractSubstrateToken();
    if (!afterToken) {
      return err(createError(
        ErrorCode.AUTH_EXPIRED,
        'Token refresh failed - no token found after refresh attempt.',
        { suggestions: ['Call teams_login to re-authenticate'] }
      ));
    }

    const newExpiry = afterToken.expiry;
    const minutesGained = Math.round(
      (newExpiry.getTime() - previousExpiry.getTime()) / 1000 / 60
    );

    // Check if the token was close to expiry and needed refresh
    const wasCloseToExpiry = previousExpiry.getTime() - Date.now() < TOKEN_REFRESH_THRESHOLD_MS;

    // If we needed a refresh but didn't get one, that's an error
    if (wasCloseToExpiry && newExpiry.getTime() <= previousExpiry.getTime()) {
      return err(createError(
        ErrorCode.AUTH_EXPIRED,
        'Token was not refreshed despite being close to expiry. Session may need re-authentication.',
        { suggestions: ['Call teams_login to re-authenticate'] }
      ));
    }

    return ok({
      newExpiry,
      previousExpiry,
      minutesGained,
      refreshNeeded: wasCloseToExpiry,
    });

  } catch (error) {
    // Clean up browser if still open
    if (manager) {
      try {
        await closeBrowser(manager, false);
      } catch {
        // Ignore cleanup errors
      }
    }

    const message = error instanceof Error ? error.message : 'Unknown error';
    return err(createError(
      ErrorCode.UNKNOWN,
      `Token refresh via browser failed: ${message}`,
      { suggestions: ['Call teams_login to re-authenticate'] }
    ));
  } finally {
    refreshInProgress = false;
  }
}
