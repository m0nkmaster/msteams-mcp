/**
 * Authentication handling for Microsoft Teams.
 * Manages login detection and manual authentication flows.
 */

import type { Page, BrowserContext } from 'playwright';
import { saveSessionState } from './context.js';
import {
  OVERLAY_STEP_PAUSE_MS,
  OVERLAY_COMPLETE_PAUSE_MS,
} from '../constants.js';
import { extractSubstrateToken, clearTokenCache } from '../auth/token-extractor.js';

/**
 * Default Teams URL for initial login.
 * 
 * For commercial tenants, this is teams.microsoft.com.
 * For GCC/GCC-High/DoD tenants, Microsoft's login flow will redirect users
 * to the appropriate URL (teams.microsoft.us, etc.) after authentication.
 * We then extract the correct base URL from DISCOVER-REGION-GTM for all API calls.
 */
const TEAMS_URL = 'https://teams.microsoft.com';

// ─────────────────────────────────────────────────────────────────────────────
// Progress Overlay UI
// ─────────────────────────────────────────────────────────────────────────────

const PROGRESS_OVERLAY_ID = 'mcp-login-progress-overlay';

/** Phases for the login progress overlay. */
type OverlayPhase = 'signed-in' | 'acquiring' | 'saving' | 'complete' | 'refreshing' | 'error';

/** Content for each overlay phase. */
const OVERLAY_CONTENT: Record<OverlayPhase, { message: string; detail: string }> = {
  'signed-in': {
    message: "You're signed in!",
    detail: 'Setting up your connection to Teams...',
  },
  'acquiring': {
    message: 'Acquiring permissions...',
    detail: 'Getting access to search and messages...',
  },
  'saving': {
    message: 'Saving your session...',
    detail: "So you won't need to log in again.",
  },
  'complete': {
    message: 'All done!',
    detail: 'This window will close automatically.',
  },
  'refreshing': {
    message: 'Refreshing your session...',
    detail: 'Updating your access tokens...',
  },
  'error': {
    message: 'Something went wrong',
    detail: 'Please try again or check the console for details.',
  },
};

/** Detail messages that cycle during the acquiring/refreshing phases. */
const ACQUIRING_DETAILS = [
  'Preparing Teams connection...',
  'Navigating to search...',
  'Waiting for API response...',
  'Acquiring search permissions...',
  'Convincing Microsoft we mean well...',
  'Negotiating with the UI...',
  'Gathering auth tokens...',
  'Good things come to those who wait...',
  'Almost there...',
];

/** Interval for cycling detail messages (ms). */
const DETAIL_CYCLE_INTERVAL_MS = 3000;

/** ID for the detail text element (for cycling updates). */
const DETAIL_ELEMENT_ID = 'mcp-login-detail';

/**
 * Shows a progress overlay for a specific phase.
 * Handles injection, content, and optional pause.
 * Failures are silently ignored - the overlay is purely cosmetic.
 */
async function showLoginProgress(
  page: Page,
  phase: OverlayPhase,
  options: { pause?: boolean } = {}
): Promise<void> {
  const content = OVERLAY_CONTENT[phase];
  const isComplete = phase === 'complete';
  const isError = phase === 'error';
  const isAnimated = phase === 'acquiring' || phase === 'refreshing';

  try {
    await page.evaluate(({ id, detailId, message, detail, complete, error, animated, cycleDetails, cycleInterval }) => {
      // Remove existing overlay if present, clearing any running timer
      const existing = document.getElementById(id);
      if (existing) {
        const existingWithTimer = existing as HTMLElement & { _cycleTimer?: ReturnType<typeof setInterval> };
        if (existingWithTimer._cycleTimer) {
          clearInterval(existingWithTimer._cycleTimer);
        }
        existing.remove();
      }

      // Remove any existing style element
      const existingStyle = document.getElementById(`${id}-style`);
      if (existingStyle) {
        existingStyle.remove();
      }

      // Add keyframe animations for spinner
      const style = document.createElement('style');
      style.id = `${id}-style`;
      style.textContent = `
        @keyframes mcp-spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
        @keyframes mcp-pulse {
          0%, 100% { opacity: 1; }
          50% { opacity: 0.6; }
        }
        @keyframes mcp-fade {
          0% { opacity: 0; transform: translateY(4px); }
          15% { opacity: 1; transform: translateY(0); }
          85% { opacity: 1; transform: translateY(0); }
          100% { opacity: 0; transform: translateY(-4px); }
        }
      `;
      document.head.appendChild(style);

      // Create overlay container
      const overlay = document.createElement('div');
      overlay.id = id;
      Object.assign(overlay.style, {
        position: 'fixed',
        top: '0',
        left: '0',
        right: '0',
        bottom: '0',
        background: 'rgba(0, 0, 0, 0.7)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        zIndex: '999999',
        fontFamily: "'Segoe UI', system-ui, sans-serif",
      });

      // Create modal card
      const modal = document.createElement('div');
      Object.assign(modal.style, {
        background: 'white',
        borderRadius: '12px',
        padding: '40px 48px',
        maxWidth: '420px',
        textAlign: 'center',
        boxShadow: '0 8px 32px rgba(0, 0, 0, 0.3)',
      });

      // Create icon container (for animation)
      const iconContainer = document.createElement('div');
      Object.assign(iconContainer.style, {
        width: '64px',
        height: '64px',
        margin: '0 auto 24px',
        position: 'relative',
      });

      // Create icon
      const icon = document.createElement('div');
      const iconBg = error ? '#c42b1c' : complete ? '#107c10' : '#5b5fc7';
      Object.assign(icon.style, {
        width: '64px',
        height: '64px',
        borderRadius: '50%',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        fontSize: '32px',
        background: iconBg,
        color: 'white',
      });

      if (animated) {
        // Spinner ring for animated states
        const spinner = document.createElement('div');
        Object.assign(spinner.style, {
          position: 'absolute',
          top: '-4px',
          left: '-4px',
          width: '72px',
          height: '72px',
          borderRadius: '50%',
          border: '3px solid transparent',
          borderTopColor: iconBg,
          borderRightColor: iconBg,
          animation: 'mcp-spin 1.2s linear infinite',
        });
        iconContainer.appendChild(spinner);
        icon.textContent = '⋯';
        icon.style.animation = 'mcp-pulse 2s ease-in-out infinite';
      } else {
        icon.textContent = error ? '✕' : complete ? '✓' : '⋯';
      }

      iconContainer.appendChild(icon);

      // Create title
      const title = document.createElement('h2');
      Object.assign(title.style, {
        margin: '0 0 12px',
        fontSize: '20px',
        fontWeight: '600',
        color: '#242424',
      });
      title.textContent = message;

      // Create detail text
      const detailEl = document.createElement('p');
      detailEl.id = detailId;
      Object.assign(detailEl.style, {
        margin: '0',
        fontSize: '14px',
        color: '#616161',
        lineHeight: '1.5',
        minHeight: '21px', // Prevent layout shift
      });
      if (animated) {
        detailEl.style.animation = `mcp-fade ${cycleInterval}ms ease-in-out infinite`;
      }
      detailEl.textContent = detail;

      // Assemble and append
      modal.appendChild(iconContainer);
      modal.appendChild(title);
      modal.appendChild(detailEl);
      overlay.appendChild(modal);
      document.body.appendChild(overlay);

      // Set up detail cycling for animated states
      if (animated && cycleDetails && cycleDetails.length > 0) {
        let detailIndex = 0;
        const cycleTimer = setInterval(() => {
          const el = document.getElementById(detailId);
          if (el) {
            el.textContent = cycleDetails[detailIndex];
            detailIndex = (detailIndex + 1) % cycleDetails.length;
          } else {
            clearInterval(cycleTimer);
          }
        }, cycleInterval);

        // Store timer ID on overlay for potential future cleanup
        (overlay as HTMLElement & { _cycleTimer?: ReturnType<typeof setInterval> })._cycleTimer = cycleTimer;
      }
    }, {
      id: PROGRESS_OVERLAY_ID,
      detailId: DETAIL_ELEMENT_ID,
      message: content.message,
      detail: content.detail,
      complete: isComplete,
      error: isError,
      animated: isAnimated,
      cycleDetails: isAnimated ? ACQUIRING_DETAILS : [],
      cycleInterval: DETAIL_CYCLE_INTERVAL_MS,
    });

    // Pause if requested (for steps that need user to see the message)
    if (options.pause) {
      const pauseMs = isComplete ? OVERLAY_COMPLETE_PAUSE_MS : OVERLAY_STEP_PAUSE_MS;
      await page.waitForTimeout(pauseMs);
    }
  } catch {
    // Overlay is cosmetic - don't fail login if it can't be shown
    // To debug: change to `catch (e)` and add `console.debug('[overlay]', e);`
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Authentication Detection
// ─────────────────────────────────────────────────────────────────────────────

// URLs that indicate we're in a login flow
const LOGIN_URL_PATTERNS = [
  'login.microsoftonline.com',
  'login.live.com',
  'login.microsoft.com',
];

// Selectors that indicate successful authentication
const AUTH_SUCCESS_SELECTORS = [
  '[data-tid="app-bar"]',
  '[data-tid="search-box"]',
  'input[placeholder*="Search"]',
  '[data-tid="chat-list"]',
  '[data-tid="team-list"]',
];

export interface AuthStatus {
  isAuthenticated: boolean;
  isOnLoginPage: boolean;
  currentUrl: string;
}

/**
 * Checks if the current page URL indicates a login flow.
 */
function isLoginUrl(url: string): boolean {
  return LOGIN_URL_PATTERNS.some(pattern => url.includes(pattern));
}

/**
 * Checks if the page shows authenticated Teams content.
 */
async function hasAuthenticatedContent(page: Page): Promise<boolean> {
  for (const selector of AUTH_SUCCESS_SELECTORS) {
    try {
      const count = await page.locator(selector).count();
      if (count > 0) {
        return true;
      }
    } catch {
      // Selector not found, continue checking others
    }
  }
  return false;
}

/**
 * Waits for the Teams SPA to be fully loaded and interactive.
 * 
 * Teams is a heavy SPA — after navigation, the shell loads quickly but MSAL
 * (which manages tokens) needs time to bootstrap. We wait for key UI elements
 * that only render after the app has fully initialised, then wait for network
 * activity to settle (indicating MSAL token requests have completed).
 */
async function waitForTeamsReady(
  page: Page,
  log: (msg: string) => void,
  timeoutMs: number = 60000
): Promise<boolean> {
  log('Waiting for Teams to fully load...');

  try {
    // Wait for any SPA UI element that indicates the app has bootstrapped
    await page.waitForSelector(
      AUTH_SUCCESS_SELECTORS.join(', '),
      { timeout: timeoutMs }
    );
    log('Teams UI elements detected.');

    // Wait for network to settle — MSAL token refresh requests should complete
    try {
      await page.waitForLoadState('networkidle', { timeout: 15000 });
      log('Network settled.');
    } catch {
      // Network may not become fully idle (websockets, polling), that's OK
      log('Network did not fully settle, continuing...');
    }

    return true;
  } catch {
    log('Teams did not fully load within timeout.');
    return false;
  }
}

/**
 * Triggers MSAL to acquire the Substrate token.
 * 
 * MSAL only acquires tokens for specific scopes when the app makes API calls
 * requiring those scopes. The Substrate API is only used for search, so we
 * perform a search to trigger token acquisition.
 * 
 * Returns true if a Substrate API call was detected, false otherwise.
 */
async function triggerTokenAcquisition(
  page: Page,
  log: (msg: string) => void
): Promise<boolean> {
  log('Triggering token acquisition...');

  try {
    // Wait for Teams SPA to be fully loaded (MSAL must bootstrap first)
    const ready = await waitForTeamsReady(page, log);
    if (!ready) {
      log('Teams did not load — cannot trigger token acquisition.');
      return false;
    }

    // Set up Substrate API listener BEFORE triggering search
    let substrateDetected = false;
    const substratePromise = page.waitForResponse(
      resp => resp.url().includes('substrate.office.com') && resp.status() === 200,
      { timeout: 30000 }
    ).then(() => {
      substrateDetected = true;
    }).catch(() => {
      // Timeout — no Substrate call detected
    });

    // Try multiple methods to trigger search
    let searchTriggered = false;

    // Method 1: Navigate to search results URL (triggers Substrate API call directly)
    log('Navigating to search results...');
    try {
      await page.goto('https://teams.microsoft.com/v2/#/search?query=test', {
        waitUntil: 'domcontentloaded',
        timeout: 30000,
      });
      searchTriggered = true;
      log('Search results page loaded.');
    } catch (e) {
      log(`Search navigation failed: ${e instanceof Error ? e.message : String(e)}`);
    }

    // Method 2: Fallback - focus and type
    if (!searchTriggered) {
      log('Trying focus+type fallback...');
      try {
        const focused = await page.evaluate(() => {
          const selectors = [
            '#ms-searchux-input',
            '[data-tid="searchInputField"]',
            'input[placeholder*="Search"]',
          ];
          for (const sel of selectors) {
            const el = document.querySelector(sel) as HTMLInputElement | null;
            if (el) {
              el.focus();
              el.click();
              return true;
            }
          }
          return false;
        });

        if (focused) {
          await page.waitForTimeout(500);
          await page.keyboard.type('test', { delay: 30 });
          await page.keyboard.press('Enter');
          searchTriggered = true;
          log('Search submitted via typing.');
        }
      } catch {
        // Continue
      }
    }

    // Method 3: Keyboard shortcut fallback
    if (!searchTriggered) {
      log('Trying keyboard shortcut...');
      const isMac = process.platform === 'darwin';
      await page.keyboard.press(isMac ? 'Meta+e' : 'Control+e');
      await page.waitForTimeout(1000);
      await page.keyboard.type('is:Messages', { delay: 30 });
      await page.keyboard.press('Enter');
      searchTriggered = true;
    }

    // Wait for the Substrate API response
    log('Waiting for Substrate API...');
    await substratePromise;

    if (substrateDetected) {
      log('Substrate API call detected — tokens acquired.');
    } else {
      log('No Substrate API call detected within timeout.');
    }

    // Give MSAL a moment to persist tokens to localStorage
    await page.waitForTimeout(2000);

    // Close search and reset
    try {
      await page.keyboard.press('Escape');
      await page.waitForTimeout(500);
    } catch {
      // Page may have navigated, ignore
    }

    log('Token acquisition complete.');
    return substrateDetected;
  } catch (error) {
    log(`Token acquisition warning: ${error instanceof Error ? error.message : String(error)}`);
    return false;
  }
}

/**
 * Gets the current authentication status.
 */
export async function getAuthStatus(page: Page): Promise<AuthStatus> {
  const currentUrl = page.url();
  const onLoginPage = isLoginUrl(currentUrl);

  // If on login page, definitely not authenticated
  if (onLoginPage) {
    return {
      isAuthenticated: false,
      isOnLoginPage: true,
      currentUrl,
    };
  }

  // If on Teams domain, check for authenticated content
  if (currentUrl.includes('teams.microsoft.com')) {
    const hasContent = await hasAuthenticatedContent(page);
    return {
      isAuthenticated: hasContent,
      isOnLoginPage: false,
      currentUrl,
    };
  }

  // Unknown state
  return {
    isAuthenticated: false,
    isOnLoginPage: false,
    currentUrl,
  };
}

/** Timeout for detecting login redirect (ms). */
const LOGIN_REDIRECT_TIMEOUT_MS = 5000;

/** URL patterns that indicate we're on a Teams page (not redirected elsewhere). */
const TEAMS_URL_PATTERNS = [
  'teams.microsoft.com',
  'teams.microsoft.us',      // GCC-High
  'dod.teams.microsoft.us',  // DoD
  'teams.cloud.microsoft',   // New Teams URL
];

/**
 * Checks if a URL is a Teams domain.
 */
function isTeamsUrl(url: string): boolean {
  return TEAMS_URL_PATTERNS.some(pattern => url.includes(pattern));
}

/**
 * Navigates to Teams and checks authentication status.
 * 
 * Uses a fast redirect-based detection: if we're not redirected to a login
 * page within a few seconds, the session is valid. This is much faster than
 * waiting for the full Teams SPA to render (which can take 30+ seconds).
 * 
 * Returns isAuthenticated: false if we can't confirm we're on Teams, to avoid
 * silently failing with an invisible browser stuck on an unexpected page.
 */
export async function navigateToTeams(page: Page): Promise<AuthStatus> {
  // Set up a promise that resolves when we detect a login redirect
  let redirectDetected = false;
  
  // Handler for detecting login redirects
  const handleFrameNavigated = (frame: import('playwright').Frame) => {
    if (frame === page.mainFrame() && isLoginUrl(frame.url())) {
      redirectDetected = true;
    }
  };

  // Listen for navigation events
  page.on('framenavigated', handleFrameNavigated);

  try {
    // Navigate to Teams
    await page.goto(TEAMS_URL, { waitUntil: 'domcontentloaded' });

    // Wait for either:
    // 1. A redirect to login page (detected via framenavigated)
    // 2. Timeout expires (no redirect = session valid)
    // 
    // Research shows login redirect happens ~3-4 seconds after navigation
    // when session is invalid (MSAL tries silent auth first, then redirects).
    // 5 seconds gives enough buffer while still being fast.
    const startTime = Date.now();
    while (Date.now() - startTime < LOGIN_REDIRECT_TIMEOUT_MS) {
      if (redirectDetected) break;
      await page.waitForTimeout(100); // Check every 100ms
    }
  } finally {
    // Clean up listener to avoid memory leaks
    page.off('framenavigated', handleFrameNavigated);
  }

  // Check final state
  const currentUrl = page.url();
  
  // Definitely on login page
  if (redirectDetected || isLoginUrl(currentUrl)) {
    return {
      isAuthenticated: false,
      isOnLoginPage: true,
      currentUrl,
    };
  }

  // Verify we're actually on a Teams page (not some unexpected redirect)
  // If we ended up somewhere unexpected, treat as unauthenticated to avoid
  // silently failing with a headless browser stuck on the wrong page
  if (!isTeamsUrl(currentUrl)) {
    return {
      isAuthenticated: false,
      isOnLoginPage: false,  // Not on login, but also not on Teams
      currentUrl,
    };
  }

  // On a Teams URL and no redirect to login = session is valid
  return {
    isAuthenticated: true,
    isOnLoginPage: false,
    currentUrl,
  };
}

/**
 * Waits for the user to complete manual authentication.
 * Returns when authenticated or throws after timeout.
 *
 * @param page - The page to monitor
 * @param context - Browser context for saving session
 * @param timeoutMs - Maximum time to wait (default: 5 minutes)
 * @param onProgress - Callback for progress updates
 * @param showOverlay - Whether to show progress overlay (default: true for visible browsers)
 */
export async function waitForManualLogin(
  page: Page,
  context: BrowserContext,
  timeoutMs: number = 5 * 60 * 1000,
  onProgress?: (message: string) => void,
  showOverlay: boolean = true
): Promise<void> {
  const startTime = Date.now();
  const log = onProgress ?? console.log;

  log('Waiting for manual login...');

  while (Date.now() - startTime < timeoutMs) {
    const status = await getAuthStatus(page);

    if (status.isAuthenticated) {
      log('Authentication successful!');

      // Show progress through login steps (only if overlay enabled)
      if (showOverlay) {
        await showLoginProgress(page, 'signed-in', { pause: true });
        await showLoginProgress(page, 'acquiring');
      }

      // Wait for Teams SPA to fully load, then trigger search for token acquisition
      await waitForTeamsReady(page, log);
      const acquired = await triggerTokenAcquisition(page, log);

      if (showOverlay) {
        await showLoginProgress(page, 'saving');
      }

      // Save the session state with fresh tokens
      await saveSessionState(context);
      clearTokenCache();
      log('Session state saved.');

      // Verify tokens were actually acquired
      const token = extractSubstrateToken();
      if (!token || token.expiry.getTime() <= Date.now()) {
        log(`Warning: No valid Substrate token after login (Substrate API detected: ${acquired}).`);
        if (showOverlay) {
          await showLoginProgress(page, 'error', { pause: true });
        }
        throw new Error('Login completed but token acquisition failed. Please try again.');
      }

      if (showOverlay) {
        await showLoginProgress(page, 'complete', { pause: true });
      }

      return;
    }

    // Check every 2 seconds
    await page.waitForTimeout(2000);
  }

  // Show error overlay before throwing (only if overlay enabled)
  if (showOverlay) {
    await showLoginProgress(page, 'error', { pause: true });
  }

  throw new Error('Authentication timeout: user did not complete login within the allowed time');
}

/**
 * Performs a full authentication flow:
 * 1. Navigate to Teams
 * 2. Check if already authenticated
 * 3. If not, wait for manual login (or throw if headless)
 *
 * @param page - The page to use
 * @param context - Browser context for session management
 * @param onProgress - Callback for progress updates
 * @param showOverlay - Whether to show progress overlay (default: true for visible browsers)
 * @param headless - If true, throw immediately if user interaction is required (default: false)
 */
export async function ensureAuthenticated(
  page: Page,
  context: BrowserContext,
  onProgress?: (message: string) => void,
  showOverlay: boolean = true,
  headless: boolean = false
): Promise<void> {
  const log = onProgress ?? console.log;

  log('Navigating to Teams...');
  const status = await navigateToTeams(page);

  if (status.isAuthenticated) {
    log('Already authenticated.');

    if (showOverlay) {
      await showLoginProgress(page, 'refreshing');
    }

    // Wait for Teams SPA to fully load, then trigger search for token acquisition
    const acquired = await triggerTokenAcquisition(page, log);

    if (showOverlay) {
      await showLoginProgress(page, 'saving');
    }

    // Save the session state with fresh tokens
    await saveSessionState(context);

    // Clear token cache so we re-extract from the freshly saved session
    clearTokenCache();

    // Verify tokens were actually refreshed
    const afterToken = extractSubstrateToken();
    const tokenIsValid = afterToken !== null && afterToken.expiry.getTime() > Date.now();

    if (!tokenIsValid) {
      const msg = `Token refresh failed: no valid token after refresh (Substrate API detected: ${acquired}).`;
      log(msg);
      if (showOverlay) {
        await showLoginProgress(page, 'error', { pause: true });
      }
      throw new Error(msg + ' Please use teams_login with forceNew to re-authenticate.');
    }

    if (showOverlay) {
      await showLoginProgress(page, 'complete', { pause: true });
    }

    return;
  }

  // User interaction required - fail fast if headless
  if (headless) {
    const reason = status.isOnLoginPage 
      ? 'Login page detected - user credentials required'
      : `Unexpected page state: ${status.currentUrl}`;
    throw new Error(`Headless SSO failed: ${reason}`);
  }

  if (status.isOnLoginPage) {
    log('Login required. Please complete authentication in the browser window.');
    await waitForManualLogin(page, context, undefined, onProgress, showOverlay);

    // Navigate back to Teams after login (in case we're on a callback URL)
    await navigateToTeams(page);
  } else {
    // Unexpected state - might need manual intervention
    log('Unexpected page state. Waiting for authentication...');
    await waitForManualLogin(page, context, undefined, onProgress, showOverlay);
  }
}

/**
 * Forces a new login by clearing session and navigating to Teams.
 */
export async function forceNewLogin(
  page: Page,
  context: BrowserContext,
  onProgress?: (message: string) => void
): Promise<void> {
  const log = onProgress ?? console.log;

  log('Starting fresh login...');

  // Clear cookies to force re-authentication
  await context.clearCookies();

  // Navigate and wait for login
  await navigateToTeams(page);
  await waitForManualLogin(page, context, undefined, onProgress);
}
