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
} from '../auth/session-store.js';
import { clearRegionCache } from '../utils/auth-guards.js';
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

/**
 * Path to the Chromium SingletonLock file.
 * This file prevents multiple processes from using the same profile.
 */
const SINGLETON_LOCK_PATH = path.join(BROWSER_PROFILE_DIR, 'SingletonLock');

/**
 * Checks if a process with the given PID is running.
 */
function isProcessRunning(pid: number): boolean {
  try {
    // kill with signal 0 checks if process exists without killing it
    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
}

/**
 * Cleans up stale SingletonLock files left by crashed browser sessions.
 * 
 * On Unix systems, SingletonLock is a symlink whose target contains the PID.
 * On Windows, it's a regular file. If the owning process is no longer running,
 * we can safely remove the lock.
 * 
 * @returns true if a stale lock was removed, false otherwise
 */
function cleanupStaleSingletonLock(): boolean {
  if (!fs.existsSync(SINGLETON_LOCK_PATH)) {
    return false;
  }

  try {
    // On Unix, SingletonLock is a symlink like "hostname-12345"
    // The number after the last dash is the PID
    const linkTarget = fs.readlinkSync(SINGLETON_LOCK_PATH);
    const pidMatch = linkTarget.match(/-(\d+)$/);
    
    if (pidMatch) {
      const pid = parseInt(pidMatch[1], 10);
      
      if (isProcessRunning(pid)) {
        // Process is still running - lock is valid
        return false;
      }
      
      // Process is dead - remove stale lock
      log.info('browser', `Removing stale SingletonLock (PID ${pid} not running)`);
      fs.unlinkSync(SINGLETON_LOCK_PATH);
      return true;
    }
  } catch {
    // On Windows or if readlink fails, try checking file age
    // If the lock file is old and we can't verify the process, remove it
    try {
      const stats = fs.lstatSync(SINGLETON_LOCK_PATH);
      const ageMs = Date.now() - stats.mtimeMs;
      const oneHourMs = 60 * 60 * 1000;
      
      if (ageMs > oneHourMs) {
        log.info('browser', `Removing old SingletonLock (${Math.round(ageMs / 60000)} minutes old)`);
        fs.unlinkSync(SINGLETON_LOCK_PATH);
        return true;
      }
    } catch {
      // Can't read stats, leave it alone
    }
  }

  return false;
}

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
  { headless = true }: { headless?: boolean } = {}
): Promise<BrowserManager> {
  ensureConfigDir();

  const browser = getBrowserChoice();
  const isFirefox = browser === 'firefox';

  // Clean up any stale lock files before attempting to launch
  if (!isFirefox) {
    cleanupStaleSingletonLock();
  }

  const launchBrowser = async (): Promise<BrowserManager> => {
    const options = {
      headless,
      viewport: { width: 1280, height: 800 },
      acceptDownloads: false,
    };
    const context = isFirefox
      ? await firefox.launchPersistentContext(FIREFOX_PROFILE_DIR, options)
      : await chromium.launchPersistentContext(BROWSER_PROFILE_DIR, { ...options, channel: browser });

    // Persistent contexts start with one page; use it or create one
    const page = context.pages()[0] ?? await context.newPage();

    return { context, page };
  };

  try {
    return await launchBrowser();
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    
    // Check if this is a profile lock error and try to recover
    if (!isFirefox && (errorMessage.includes('ProcessSingleton') || errorMessage.includes('SingletonLock'))) {
      log.warn('browser', 'Profile lock detected, attempting to clean up and retry...');
      
      // Force remove the lock file and retry once
      try {
        if (fs.existsSync(SINGLETON_LOCK_PATH)) {
          fs.unlinkSync(SINGLETON_LOCK_PATH);
          log.info('browser', 'Removed SingletonLock file, retrying browser launch...');
          return await launchBrowser();
        }
      } catch (cleanupError) {
        log.error('browser', `Failed to clean up SingletonLock: ${cleanupError instanceof Error ? cleanupError.message : String(cleanupError)}`);
      }
    }

    throw new Error(
      `Could not launch ${BROWSER_NAMES[browser]}. ${INSTALL_HINTS[browser]}\n\n` +
      `Original error: ${errorMessage}`
    );
  }
}

/**
 * Saves the current browser context's session state.
 */
export async function saveSessionState(context: BrowserContext): Promise<void> {
  const state = await context.storageState();
  // MSAL.js v4+ encrypts its localStorage cache; store plaintext entries so
  // token extraction and HTTP refresh can read them. Entries that cannot be
  // decrypted are kept as-is.
  const stats = decryptStorageState(state);
  if (stats.encrypted > 0) {
    log.debug('browser', `MSAL cache entries: ${stats.encrypted} encrypted, ${stats.decrypted} decrypted`);
  }
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
