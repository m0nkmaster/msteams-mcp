/**
 * Playwright browser context management.
 * Creates and manages browser contexts with session persistence.
 * 
 * Uses the system's installed Chrome or Edge browser rather than downloading
 * Playwright's bundled Chromium. This significantly reduces install size.
 */

import { chromium, type Browser, type BrowserContext, type Page } from 'playwright';
import {
  ensureUserDataDir,
  hasSessionState,
  SESSION_STATE_PATH,
  isSessionLikelyExpired,
  writeSessionState,
  readSessionState,
} from '../auth/session-store.js';
import { clearRegionCache } from '../utils/auth-guards.js';

export interface BrowserManager {
  browser: Browser;
  context: BrowserContext;
  page: Page;
  isNewSession: boolean;
}

export interface CreateBrowserOptions {
  headless?: boolean;
  viewport?: { width: number; height: number };
}

const DEFAULT_OPTIONS: Required<CreateBrowserOptions> = {
  headless: true,
  viewport: { width: 1280, height: 800 },
};

/**
 * Determines the browser channel to use based on the platform.
 * - Windows: Use Microsoft Edge (always installed on Windows 10+)
 * - macOS/Linux: Use Chrome
 * 
 * @returns The browser channel name for Playwright
 */
function getBrowserChannel(): 'msedge' | 'chrome' {
  return process.platform === 'win32' ? 'msedge' : 'chrome';
}

/**
 * Creates a browser context with optional session state restoration.
 *
 * Uses the system's installed Chrome or Edge browser rather than downloading
 * Playwright's bundled Chromium (~180MB savings).
 *
 * @param options - Browser configuration options
 * @returns Browser manager with browser, context, and page
 * @throws Error if system browser is not found (with helpful suggestions)
 */
export async function createBrowserContext(
  options: CreateBrowserOptions = {}
): Promise<BrowserManager> {
  const opts = { ...DEFAULT_OPTIONS, ...options };

  ensureUserDataDir();

  const channel = getBrowserChannel();
  let browser: Browser;

  try {
    browser = await chromium.launch({
      headless: opts.headless,
      channel,
    });
  } catch (error) {
    const browserName = channel === 'msedge' ? 'Microsoft Edge' : 'Google Chrome';
    const installHint = channel === 'msedge'
      ? 'Edge should be pre-installed on Windows. Try updating Windows or reinstalling Edge.'
      : 'Install Chrome from https://www.google.com/chrome/ or run: npx playwright install chromium';
    
    throw new Error(
      `Could not launch ${browserName}. ${installHint}\n\n` +
      `Original error: ${error instanceof Error ? error.message : String(error)}`
    );
  }

  const hasSession = hasSessionState();
  const sessionExpired = isSessionLikelyExpired();

  // Restore session if we have one and it's not ancient
  const shouldRestoreSession = hasSession && !sessionExpired;

  let context: BrowserContext;

  if (shouldRestoreSession) {
    try {
      // Read the decrypted session state
      const state = readSessionState();
      if (state) {
        // Create a temporary file for Playwright (it needs a file path)
        // We write the decrypted state to a temp location
        const tempPath = SESSION_STATE_PATH + '.tmp';
        const fs = await import('fs');
        fs.writeFileSync(tempPath, JSON.stringify(state), { mode: 0o600 });

        try {
          context = await browser.newContext({
            storageState: tempPath,
            viewport: opts.viewport,
          });
        } finally {
          // Clean up temp file
          fs.unlinkSync(tempPath);
        }
      } else {
        throw new Error('Failed to read session state');
      }
    } catch (error) {
      console.warn('Failed to restore session state, starting fresh:', error);
      context = await browser.newContext({
        viewport: opts.viewport,
      });
    }
  } else {
    context = await browser.newContext({
      viewport: opts.viewport,
    });
  }

  const page = await context.newPage();

  return {
    browser,
    context,
    page,
    isNewSession: !shouldRestoreSession,
  };
}

/**
 * Saves the current browser context's session state.
 */
export async function saveSessionState(context: BrowserContext): Promise<void> {
  const state = await context.storageState();
  writeSessionState(state);
  // Clear cached region config so new session values are picked up
  clearRegionCache();
}

/**
 * Closes the browser and optionally saves session state.
 */
export async function closeBrowser(
  manager: BrowserManager,
  saveSession: boolean = true
): Promise<void> {
  if (saveSession) {
    await saveSessionState(manager.context);
  }
  await manager.context.close();
  await manager.browser.close();
}
