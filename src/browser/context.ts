/**
 * Playwright browser context management.
 * Creates and manages browser contexts with session persistence.
 * 
 * Uses the system's installed Chrome or Edge browser rather than downloading
 * Playwright's bundled Chromium. This significantly reduces install size.
 * Set TEAMS_BROWSER=firefox to use Playwright's Firefox build instead.
 * 
 * All modes (headless and visible) use a persistent browser profile stored at
 * ~/.teams-mcp-server/browser-profile/ (browser-profile-firefox/ for Firefox).
 * This means:
 * - Microsoft session cookies persist across launches (longer-lived than MSAL tokens)
 * - Headless token refresh can silently re-authenticate using the profile's session
 * - Visible login retains extensions (e.g. Bitwarden) and form autofill data
 * - No need for storageState temp files or encrypted session restoration for browser use
 */

import { chromium, firefox, type BrowserContext, type Page } from 'playwright';
import * as fs from 'fs';
import * as path from 'path';
import {
  ensureConfigDir,
  CONFIG_DIR,
  writeSessionState,
  clearTokenCache,
  readSessionState,
} from '../auth/session-store.js';
import { clearRegionCache } from '../utils/auth-guards.js';
import { extractSubstrateToken, extractSkypeSpacesToken } from '../auth/token-extractor.js';
import { decryptStorageState } from '../auth/msal-cache-crypto.js';
import * as log from '../utils/logger.js';

export interface BrowserManager {
  context: BrowserContext;
  page: Page;
}

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

/** Firefox needs its own profile directory: the on-disk format differs from Chromium's. */
const FIREFOX_PROFILE_DIR = path.join(CONFIG_DIR, 'browser-profile-firefox');

const recoveryProfiles = new WeakMap<BrowserContext, { name: string; discardOnClose: boolean }>();

export type BrowserChoice = 'chrome' | 'msedge' | 'firefox';

const BROWSER_CHOICES: readonly BrowserChoice[] = ['chrome', 'msedge', 'firefox'];

/**
 * Determines which browser to launch.
 * - TEAMS_BROWSER (chrome, msedge or firefox) overrides the default.
 * - Otherwise Edge on Windows (always installed on Windows 10+), Chrome elsewhere.
 */
export function getBrowserChoice(): BrowserChoice {
  const requested = process.env.TEAMS_BROWSER?.trim().toLowerCase();
  if (requested) {
    if ((BROWSER_CHOICES as readonly string[]).includes(requested)) {
      return requested as BrowserChoice;
    }
    log.warn('browser', `Ignoring unknown TEAMS_BROWSER "${requested}"; expected one of ${BROWSER_CHOICES.join(', ')}`);
  }
  return process.platform === 'win32' ? 'msedge' : 'chrome';
}

const BROWSER_NAMES: Record<BrowserChoice, string> = {
  chrome: 'Google Chrome',
  msedge: 'Microsoft Edge',
  firefox: 'Firefox',
};

const INSTALL_HINTS: Record<BrowserChoice, string> = {
  chrome: 'Install Chrome from https://www.google.com/chrome/, or set TEAMS_BROWSER=firefox and run: npx playwright install firefox',
  msedge: 'Edge should be pre-installed on Windows. Try updating Windows or reinstalling Edge.',
  firefox: 'Run: npx playwright install firefox',
};

/**
 * Creates a browser context using a persistent profile.
 *
 * Chrome and Edge use the system's installed browser rather than downloading
 * Playwright's bundled Chromium (~180MB savings). Firefox uses Playwright's
 * own Firefox build, since Playwright cannot drive a stock Firefox.
 *
 * The persistent profile at ~/.teams-mcp-server/browser-profile/ (or
 * browser-profile-firefox/) is shared between headless and visible modes.
 * This provides:
 * - Silent headless re-auth via long-lived Microsoft session cookies
 * - Extensions (e.g. Bitwarden) and form autofill for visible login
 * - No storageState temp file management needed
 *
 * Note: Only one process can use the profile at a time (profile lock).
 * The MCP server serialises tool calls, and token-refresh checks for an active
 * browser before attempting refresh to avoid lock contention.
 *
 * @param options.headless - Run without a visible window (default: true)
 * @returns Browser manager with context and page
 * @throws Error if the browser is not found (with helpful suggestions)
 */
export async function createBrowserContext(
  { headless = true, freshProfile = false }: { headless?: boolean; freshProfile?: boolean } = {}
): Promise<BrowserManager> {
  ensureConfigDir();

  const browser = getBrowserChoice();
  const isFirefox = browser === 'firefox';

  const defaultDir = isFirefox ? FIREFOX_PROFILE_DIR : BROWSER_PROFILE_DIR;
  const prefix = `recovery-${browser}-`;
  const savedProfile = readSessionState()?.browserProfile;
  let profileDir = defaultDir;
  if (typeof savedProfile === 'string' && new RegExp(`^${prefix}[a-zA-Z0-9]{6}$`).test(savedProfile)) {
    profileDir = path.join(CONFIG_DIR, savedProfile);
  }
  if (freshProfile) profileDir = fs.mkdtempSync(path.join(CONFIG_DIR, prefix));

  try {
    const options = {
      headless,
      viewport: { width: 1280, height: 800 },
      acceptDownloads: false,
    };
    const context = isFirefox
      ? await firefox.launchPersistentContext(profileDir, options)
      : await chromium.launchPersistentContext(profileDir, { ...options, channel: browser });

    // Persistent contexts start with one page; use it or create one
    const page = context.pages()[0] ?? await context.newPage();

    if (profileDir !== defaultDir) {
      recoveryProfiles.set(context, { name: path.basename(profileDir), discardOnClose: freshProfile });
    }
    return { context, page };
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    
    if (freshProfile) fs.rmSync(profileDir, { recursive: true, force: true });
    throw new Error(
      `Could not launch ${BROWSER_NAMES[browser]}. ${INSTALL_HINTS[browser]}\n\n` +
      `Original error: ${errorMessage}`
    );
  }
}

/**
 * Saves the current browser context's session state.
 */
export async function saveAuthenticatedSession(context: BrowserContext, allowPartial: boolean): Promise<boolean> {
  const state = await context.storageState();
  // MSAL.js v4+ encrypts its localStorage cache; store plaintext entries so
  // token extraction and HTTP refresh can read them. Entries that cannot be
  // decrypted are kept as-is.
  const stats = decryptStorageState(state);
  if (stats.encrypted > 0) {
    log.debug('browser', `MSAL cache entries: ${stats.encrypted} encrypted, ${stats.decrypted} decrypted`);
  }
  if (!extractSubstrateToken(state) && !(allowPartial && extractSkypeSpacesToken(state))) return false;
  const profile = recoveryProfiles.get(context);
  writeSessionState({ ...state, browserProfile: profile?.name });
  if (profile) profile.discardOnClose = false;
  clearTokenCache();
  // Clear cached region config so new session values are picked up
  clearRegionCache();
  return true;
}

export async function closeBrowser(manager: BrowserManager): Promise<void> {
  await manager.context.close();
  const profile = recoveryProfiles.get(manager.context);
  if (profile?.discardOnClose) fs.rmSync(path.join(CONFIG_DIR, profile.name), { recursive: true, force: true });
}
