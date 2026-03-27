/**
 * Import Microsoft SSO cookies from the user's Chrome profile into a Playwright context.
 *
 * When the Teams MCP server needs to open a visible browser for login,
 * the Playwright profile is isolated from the user's real Chrome — so Microsoft
 * can't recognise the user via SSO. This module copies the relevant Microsoft
 * cookies from the user's actual Chrome work profile, enabling silent SSO in
 * the Playwright browser and eliminating the need to re-type credentials.
 *
 * macOS only (Chrome cookies are encrypted with a Keychain-backed key).
 * Fails gracefully on other platforms or when Chrome isn't available.
 *
 * Configuration:
 *   TEAMS_MCP_CHROME_PROFILE env var — Chrome profile directory name
 *   (e.g. "Profile 1"). If unset, auto-detects from Chrome's Local State.
 */

import * as crypto from 'crypto';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { execSync } from 'child_process';
import type { BrowserContext } from 'playwright';
import * as log from '../utils/logger.js';

// Microsoft domains whose cookies enable SSO
const MICROSOFT_DOMAINS = [
  '%microsoftonline%',
  '%login.live.com%',
  '%login.microsoft.com%',
  '%microsoft.com%',
  '%office.com%',
  '%office365.com%',
];

const CHROME_DATA_DIR = path.join(
  os.homedir(),
  'Library',
  'Application Support',
  'Google',
  'Chrome'
);

interface ChromeProfile {
  dirName: string;  // e.g. "Profile 1"
  name: string;     // e.g. "corp.example.com"
  gaiaName: string; // e.g. "Jane Smith"
}

interface RawCookie {
  host_key: string;
  name: string;
  encrypted_value_hex: string;
  path: string;
  expires_utc: number;  // Chrome epoch: microseconds since 1601-01-01
  is_secure: number;
  is_httponly: number;
  samesite: number;
}

// ─────────────────────────────────────────────────────────────────────────────
// Chrome Profile Detection
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Lists Chrome profiles from the Local State file.
 */
function listChromeProfiles(): ChromeProfile[] {
  const localStatePath = path.join(CHROME_DATA_DIR, 'Local State');
  if (!fs.existsSync(localStatePath)) return [];

  try {
    const localState = JSON.parse(fs.readFileSync(localStatePath, 'utf8'));
    const infoCache = localState?.profile?.info_cache;
    if (!infoCache || typeof infoCache !== 'object') return [];

    return Object.entries(infoCache).map(([dirName, info]: [string, unknown]) => {
      const i = info as Record<string, unknown>;
      return {
        dirName,
        name: String(i.name ?? ''),
        gaiaName: String(i.gaia_name ?? ''),
      };
    });
  } catch {
    return [];
  }
}

/**
 * Selects the Chrome profile to import cookies from.
 *
 * Priority:
 * 1. TEAMS_MCP_CHROME_PROFILE env var (exact dir name like "Profile 1")
 * 2. Auto-detect: first profile whose name looks like a work/corporate account
 * 3. null if no suitable profile found
 */
function selectChromeProfile(): ChromeProfile | null {
  const profiles = listChromeProfiles();
  if (profiles.length === 0) return null;

  // Priority 1: explicit env var
  const envProfile = process.env.TEAMS_MCP_CHROME_PROFILE;
  if (envProfile) {
    const match = profiles.find(p => p.dirName === envProfile);
    if (match) return match;
    log.warn('cookie-import', `TEAMS_MCP_CHROME_PROFILE="${envProfile}" not found. Available: ${profiles.map(p => `${p.dirName} (${p.name})`).join(', ')}`);
    return null;
  }

  // Priority 2: auto-detect work profile (contains a domain-like name)
  const workProfile = profiles.find(p =>
    /\.[a-z]{2,}$/i.test(p.name) || // name contains a domain (e.g. "corp.example.com")
    p.name.toLowerCase().includes('work') ||
    p.name.toLowerCase().includes('corp')
  );
  if (workProfile) return workProfile;

  // Skip auto-import if we can't identify a work profile
  return null;
}

// ─────────────────────────────────────────────────────────────────────────────
// Cookie Decryption (macOS)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Cached AES key derived from the Chrome Safe Storage Keychain password.
 * Cached for the lifetime of the MCP server process so the Keychain
 * prompt only appears once (on first cookie import after server start).
 *
 * Tip: click "Always Allow" on the macOS Keychain dialog to permanently
 * authorize this process — then no prompt appears at all.
 */
let cachedKey: Buffer | null = null;

/**
 * Gets (or retrieves from cache) the AES key for Chrome cookie decryption.
 * Accesses the macOS Keychain only on first call, caches for process lifetime.
 */
function getDecryptionKey(): Buffer | null {
  if (cachedKey) return cachedKey;

  let password: string;
  try {
    password = execSync(
      'security find-generic-password -s "Chrome Safe Storage" -w',
      { encoding: 'utf8', timeout: 5000 }
    ).trim();
  } catch {
    return null;
  }

  cachedKey = crypto.pbkdf2Sync(password, 'saltysalt', 1003, 16, 'sha1');
  return cachedKey;
}

/**
 * Decrypts a Chrome cookie value.
 * Chrome macOS cookies are prefixed with 'v10' followed by AES-128-CBC ciphertext.
 */
function decryptCookieValue(hexValue: string, key: Buffer): string | null {
  try {
    const encrypted = Buffer.from(hexValue, 'hex');

    // v10 prefix check (0x76 0x31 0x30)
    if (encrypted.length < 4 || encrypted[0] !== 0x76 || encrypted[1] !== 0x31 || encrypted[2] !== 0x30) {
      // Not encrypted or unknown format — try as plain text
      return encrypted.toString('utf8');
    }

    const ciphertext = encrypted.subarray(3);
    const iv = Buffer.alloc(16, 0x20); // 16 bytes of space (0x20)
    const decipher = crypto.createDecipheriv('aes-128-cbc', key, iv);
    decipher.setAutoPadding(true);
    const decrypted = Buffer.concat([decipher.update(ciphertext), decipher.final()]);
    return decrypted.toString('utf8');
  } catch {
    return null;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Cookie Reading (via sqlite3 CLI)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Reads Microsoft-related cookies from a Chrome profile's Cookies database.
 * Copies the DB to a temp file to avoid locking conflicts with running Chrome.
 */
function readChromeCookies(profileDir: string): RawCookie[] {
  const cookiesDb = path.join(CHROME_DATA_DIR, profileDir, 'Cookies');
  if (!fs.existsSync(cookiesDb)) return [];

  // Copy to temp to avoid lock conflicts with running Chrome
  const tmpDb = path.join(os.tmpdir(), `teams-mcp-cookies-${Date.now()}.db`);
  try {
    fs.copyFileSync(cookiesDb, tmpDb);
    // Also copy WAL/SHM if they exist (needed for consistent reads)
    for (const ext of ['-wal', '-shm']) {
      const src = cookiesDb + ext;
      if (fs.existsSync(src)) {
        fs.copyFileSync(src, tmpDb + ext);
      }
    }

    const whereClause = MICROSOFT_DOMAINS.map(d => `host_key LIKE '${d}'`).join(' OR ');
    const sql = `SELECT host_key, name, hex(encrypted_value) as ev, path, expires_utc, is_secure, is_httponly, samesite FROM cookies WHERE (${whereClause}) AND expires_utc > 0`;

    const output = execSync(
      `sqlite3 -separator '|||' "${tmpDb}" "${sql}"`,
      { encoding: 'utf8', timeout: 10000, maxBuffer: 1024 * 1024 }
    );

    return output
      .trim()
      .split('\n')
      .filter(line => line.includes('|||'))
      .map(line => {
        const [host_key, name, encrypted_value_hex, cookiePath, expires_utc, is_secure, is_httponly, samesite] = line.split('|||');
        return {
          host_key,
          name,
          encrypted_value_hex,
          path: cookiePath,
          expires_utc: parseInt(expires_utc, 10),
          is_secure: parseInt(is_secure, 10),
          is_httponly: parseInt(is_httponly, 10),
          samesite: parseInt(samesite, 10),
        };
      });
  } catch (err) {
    log.warn('cookie-import', `Failed to read Chrome cookies: ${err instanceof Error ? err.message : String(err)}`);
    return [];
  } finally {
    // Clean up temp files
    for (const f of [tmpDb, tmpDb + '-wal', tmpDb + '-shm']) {
      try { fs.unlinkSync(f); } catch { /* ignore */ }
    }
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Cookie Conversion
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Converts Chrome epoch (microseconds since 1601-01-01) to Unix epoch (seconds since 1970-01-01).
 */
function chromeEpochToUnix(chromeTimestamp: number): number {
  // Chrome epoch starts 11644473600 seconds before Unix epoch
  return Math.floor(chromeTimestamp / 1_000_000) - 11644473600;
}

/**
 * Maps Chrome's samesite integer to Playwright's string value.
 */
function chromeSameSiteToPlaywright(samesite: number): 'Strict' | 'Lax' | 'None' {
  switch (samesite) {
    case 2: return 'Strict';
    case 1: return 'Lax';
    default: return 'None';
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Public API
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Imports Microsoft SSO cookies from the user's Chrome profile into a Playwright context.
 *
 * This enables SSO in the Playwright browser so the user doesn't have to type
 * credentials when Microsoft redirects to the login page.
 *
 * Fails silently — cookie import is best-effort. If it doesn't work,
 * the user just has to log in manually as before.
 */
export async function importMicrosoftCookies(context: BrowserContext): Promise<void> {
  // Only supported on macOS
  if (process.platform !== 'darwin') {
    log.debug('cookie-import', 'Skipping cookie import (not macOS)');
    return;
  }

  // Check if Chrome is installed
  if (!fs.existsSync(CHROME_DATA_DIR)) {
    log.debug('cookie-import', 'Skipping cookie import (Chrome not found)');
    return;
  }

  const profile = selectChromeProfile();
  if (!profile) {
    log.debug('cookie-import', 'No Chrome work profile found. Set TEAMS_MCP_CHROME_PROFILE env var.');
    return;
  }

  log.info('cookie-import', `Importing Microsoft cookies from Chrome profile: ${profile.dirName} (${profile.name})`);

  // Get decryption key (cached after first access — only one Keychain prompt per server lifetime)
  const key = getDecryptionKey();
  if (!key) {
    log.warn('cookie-import',
      'Could not get Chrome Safe Storage password from Keychain. ' +
      'To fix, run: security set-generic-password-partition-list -S "apple-tool:,apple:" ' +
      '-a "Chrome" -s "Chrome Safe Storage" ~/Library/Keychains/login.keychain-db'
    );
    return;
  }

  // Read and decrypt cookies
  const rawCookies = readChromeCookies(profile.dirName);
  if (rawCookies.length === 0) {
    log.info('cookie-import', 'No Microsoft cookies found in Chrome profile');
    return;
  }

  const playwrightCookies: Array<{
    name: string;
    value: string;
    domain: string;
    path: string;
    expires: number;
    httpOnly: boolean;
    secure: boolean;
    sameSite: 'Strict' | 'Lax' | 'None';
  }> = [];

  for (const raw of rawCookies) {
    const value = decryptCookieValue(raw.encrypted_value_hex, key);
    if (!value) continue;

    const expires = chromeEpochToUnix(raw.expires_utc);
    // Skip expired cookies
    if (expires <= Math.floor(Date.now() / 1000)) continue;

    playwrightCookies.push({
      name: raw.name,
      value,
      domain: raw.host_key,
      path: raw.path,
      expires,
      httpOnly: raw.is_httponly === 1,
      secure: raw.is_secure === 1,
      sameSite: chromeSameSiteToPlaywright(raw.samesite),
    });
  }

  if (playwrightCookies.length === 0) {
    log.info('cookie-import', 'No valid Microsoft cookies to import');
    return;
  }

  try {
    await context.addCookies(playwrightCookies);
    log.info('cookie-import', `Imported ${playwrightCookies.length} Microsoft cookies from Chrome "${profile.name}" profile`);
  } catch (err) {
    log.warn('cookie-import', `Failed to inject cookies: ${err instanceof Error ? err.message : String(err)}`);
  }
}
