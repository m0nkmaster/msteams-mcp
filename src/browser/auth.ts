/**
 * Authentication handling for Microsoft Teams.
 * Manages login detection and manual authentication flows.
 */

import type { Page, BrowserContext } from 'playwright';
import { saveAuthenticatedSession } from './context.js';
import {
  OVERLAY_STEP_PAUSE_MS,
  OVERLAY_COMPLETE_PAUSE_MS,
  TEAMS_HOSTS,
  LOGIN_TIMEOUT_MS,
} from '../constants.js';
import * as logger from '../utils/logger.js';

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
type OverlayPhase = 'waiting' | 'complete' | 'error';

/** Content for each overlay phase. */
const OVERLAY_CONTENT: Record<OverlayPhase, { message: string; detail: string }> = {
  waiting: {
    message: 'Waiting for Teams credentials',
    detail: 'Complete sign-in if prompted. This window will close when Teams is ready.',
  },
  'complete': {
    message: 'All done!',
    detail: 'This window will close automatically.',
  },
  'error': {
    message: 'Something went wrong',
    detail: 'Please try again or check the console for details.',
  },
};

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

  try {
    await page.evaluate(({ id, message, detail, complete, error, waiting }) => {
      // Remove existing overlay if present
      const existing = document.getElementById(id);
      if (waiting && existing?.dataset.phase === 'waiting') return;
      existing?.remove();

      // Create overlay container
      const overlay = document.createElement('div');
      overlay.id = id;
      if (waiting) overlay.dataset.phase = 'waiting';
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

      if (waiting) {
        Object.assign(overlay.style, {
          inset: '16px 16px auto',
          pointerEvents: 'none',
          background: 'transparent',
        });
      }

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
        margin: '0 auto 24px',
      });
      icon.textContent = error ? '✕' : complete ? '✓' : '⋯';

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
      Object.assign(detailEl.style, {
        margin: '0',
        fontSize: '14px',
        color: '#616161',
        lineHeight: '1.5',
      });
      detailEl.textContent = detail;

      if (waiting) {
        icon.style.display = 'none';
        modal.style.padding = '12px 20px';
      }

      // Assemble and append
      modal.appendChild(icon);
      modal.appendChild(title);
      modal.appendChild(detailEl);
      overlay.appendChild(modal);
      document.documentElement.appendChild(overlay);
    }, {
      id: PROGRESS_OVERLAY_ID,
      message: content.message,
      detail: content.detail,
      waiting: phase === 'waiting',
      complete: isComplete,
      error: isError,
    });

    // Pause if requested (for steps that need user to see the message)
    if (options.pause) {
      const pauseMs = isComplete ? OVERLAY_COMPLETE_PAUSE_MS : OVERLAY_STEP_PAUSE_MS;
      await page.waitForTimeout(pauseMs);
    }
  } catch {
    // Overlay is cosmetic - don't fail login if it can't be shown
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Authentication Detection
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Checks if a URL is a Teams domain.
 */
export function isTeamsUrl(url: string): boolean {
  try {
    const parsed = new URL(url);
    return parsed.protocol === 'https:' && TEAMS_HOSTS.includes(parsed.host);
  } catch {
    return false;
  }
}

/** The profile reached Teams, but silent SSO did not produce usable credentials. */
export class StaleTeamsSession extends Error {}
export class TeamsLoginTimeout extends Error {}

export async function ensureAuthenticated(
  page: Page,
  context: BrowserContext,
  onProgress?: (message: string) => void,
  showOverlay = true,
  headless = false,
  deadline = Date.now() + LOGIN_TIMEOUT_MS,
): Promise<void> {
  const log = onProgress ?? ((message: string) => logger.info('auth', message));
  log('Navigating to Teams...');
  const stopAt = headless ? Math.min(deadline, Date.now() + 20_000) : deadline;
  const remaining = stopAt - Date.now();
  if (remaining > 0) {
    await page.goto(TEAMS_URL, { waitUntil: 'commit', timeout: Math.min(30_000, remaining) });
  }
  log(headless ? 'Waiting for Teams credentials...' : 'Complete sign-in in the browser if prompted.');
  while (Date.now() < stopAt) {
    if (isTeamsUrl(page.url())) {
      if (showOverlay) await showLoginProgress(page, 'waiting');
      if (await saveAuthenticatedSession(context, !headless)) {
        log('Session state saved.');
        if (showOverlay) await showLoginProgress(page, 'complete', { pause: true });
        return;
      }
    }
    await page.waitForTimeout(headless ? 1000 : 2000);
  }
  if (headless) {
    if (isTeamsUrl(page.url())) throw new StaleTeamsSession('Teams opened but silent SSO produced no live search credential.');
    throw new Error('Silent SSO did not return to Teams. Interactive sign-in is required.');
  }
  if (showOverlay && isTeamsUrl(page.url())) await showLoginProgress(page, 'error', { pause: true });
  throw new TeamsLoginTimeout('Teams sign-in timed out before Teams provided usable credentials.');
}
