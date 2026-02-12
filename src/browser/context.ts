/**
 * Playwright browser context management.
 * Creates and manages browser contexts with session persistence.
 * 
 * Uses the system's installed Chrome or Edge browser rather than downloading
 * Playwright's bundled Chromium. This significantly reduces install size.
 * 
 * All modes (headless and visible) use a persistent browser profile stored at
 * ~/.teams-mcp-server/browser-profile/. This means:
 * - Microsoft session cookies persist across launches (longer-lived than MSAL tokens)
 * - Headless token refresh can silently re-authenticate using the profile's session
 * - Visible login retains extensions (e.g. Bitwarden) and form autofill data
 * - No need for storageState temp files or encrypted session restoration for browser use
 */

import { chromium, type BrowserContext, type Page } from 'playwright';
import * as path from 'path';
import {
  ensureUserDataDir,
  CONFIG_DIR,
  writeSessionState,
} from '../auth/session-store.js';
import { clearRegionCache } from '../utils/auth-guards.js';

export interface BrowserManager {
  /** Always null — persistent contexts have no separate Browser object. */
  browser: null;
  context: BrowserContext;
  page: Page;
  isNewSession: boolean;
  /** Always true — all contexts use the persistent browser profile. */
  persistent: true;
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
 * Directory for the persistent browser profile.
 * This is a dedicated Chrome/Edge profile within the config dir, so extensions
 * (e.g. Bitwarden) and form autofill data persist across login sessions
 * without conflicting with the user's running browser instance.
 * 
 * Both headless and visible modes share this profile, so Microsoft's long-lived
 * session cookies enable silent headless re-authentication without user interaction.
 */
const BROWSER_PROFILE_DIR = path.join(CONFIG_DIR, 'browser-profile');

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
 * Creates a browser context using a persistent profile.
 *
 * Uses the system's installed Chrome or Edge browser rather than downloading
 * Playwright's bundled Chromium (~180MB savings).
 *
 * The persistent profile at ~/.teams-mcp-server/browser-profile/ is shared
 * between headless and visible modes. This provides:
 * - Silent headless re-auth via long-lived Microsoft session cookies
 * - Extensions (e.g. Bitwarden) and form autofill for visible login
 * - No storageState temp file management needed
 *
 * Note: Only one process can use the profile at a time (Chromium profile lock).
 * The MCP server serialises tool calls, and token-refresh checks for an active
 * browser before attempting refresh to avoid lock contention.
 *
 * @param options - Browser configuration options
 * @returns Browser manager with context and page
 * @throws Error if system browser is not found (with helpful suggestions)
 */
export async function createBrowserContext(
  options: CreateBrowserOptions = {}
): Promise<BrowserManager> {
  const opts = { ...DEFAULT_OPTIONS, ...options };

  ensureUserDataDir();

  const channel = getBrowserChannel();

  try {
    const context = await chromium.launchPersistentContext(BROWSER_PROFILE_DIR, {
      headless: opts.headless,
      channel,
      viewport: opts.viewport,
      acceptDownloads: false,
    });

    // Persistent contexts start with one page; use it or create one
    const page = context.pages()[0] ?? await context.newPage();

    return {
      browser: null,
      context,
      page,
      isNewSession: true,
      persistent: true,
    };
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
 * Closes the browser context and optionally saves session state.
 */
export async function closeBrowser(
  manager: BrowserManager,
  saveSession: boolean = true
): Promise<void> {
  if (saveSession) {
    await saveSessionState(manager.context);
  }
  await manager.context.close();
}
