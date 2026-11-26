/**
 * Secure session state storage.
 * 
 * Handles reading and writing session state with:
 * - Encryption at rest
 * - Restricted file permissions
 * - Automatic migration from plaintext
 * 
 * Session files are stored in a user-specific config directory (~/.teams-mcp-server/)
 * to ensure consistency regardless of how the server is invoked (npx, global install, etc.).
 */

import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { encrypt, decrypt } from './crypto.js';
import { SESSION_EXPIRY_HOURS, TEAMS_HOSTS } from '../constants.js';
import * as log from '../utils/logger.js';

/**
 * User-specific config directory for teams-mcp-server.
 * - Windows: %APPDATA%\teams-mcp-server\
 * - macOS/Linux: ~/.teams-mcp-server/
 */
export const CONFIG_DIR = process.platform === 'win32'
  ? path.join(process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming'), 'teams-mcp-server')
  : path.join(os.homedir(), '.teams-mcp-server');
export const SESSION_STATE_PATH = path.join(CONFIG_DIR, 'session-state.json');
export const TOKEN_CACHE_PATH = path.join(CONFIG_DIR, 'token-cache.json');

/** File permission mode: owner read/write only. */
const SECURE_FILE_MODE = 0o600;

/**
 * Monotonic counter bumped whenever local session data is deleted. An
 * in-flight refresh that read the old state compares the epoch across its
 * async work and skips the write when it changed — otherwise a refresh
 * abandoned at its deadline could write the wiped account's session back.
 */
let sessionEpoch = 0;

/** Current session epoch; capture before an async read-modify-write. */
export function getSessionEpoch(): number {
  return sessionEpoch;
}

/**
 * Ensures the config directory exists with secure permissions.
 */
export function ensureConfigDir(): void {
  fs.mkdirSync(CONFIG_DIR, { recursive: true, mode: 0o700 });
}

/** Session state as stored by Playwright. */
export interface SessionState {
  cookies: Array<{
    name: string;
    value: string;
    domain?: string;
    path?: string;
    expires?: number;
    httpOnly?: boolean;
    secure?: boolean;
    sameSite?: 'Strict' | 'Lax' | 'None';
  }>;
  origins: Array<{
    origin: string;
    localStorage: Array<{ name: string; value: string }>;
  }>;
}

/** Token cache structure. */
export interface TokenCache {
  substrateToken: string;
  substrateTokenExpiry: number;
  extractedAt: number;
}

/**
 * Writes data securely with encryption and file permissions.
 */
function writeSecure(filePath: string, data: unknown): void {
  const json = JSON.stringify(data, null, 2);
  const encrypted = encrypt(json);
  
  fs.writeFileSync(filePath, JSON.stringify(encrypted, null, 2), { 
    mode: SECURE_FILE_MODE,
    encoding: 'utf8',
  });
}

/**
 * Reads and decrypts data. Returns null if missing or undecryptable
 * (different machine, corrupted).
 */
function readSecure<T>(filePath: string): T | null {
  if (!fs.existsSync(filePath)) return null;
  try {
    return JSON.parse(decrypt(JSON.parse(fs.readFileSync(filePath, 'utf8')))) as T;
  } catch (error) {
    log.error('session-store', `Failed to read ${filePath}: ${error instanceof Error ? error.message : error}`);
    return null;
  }
}

/**
 * Checks if session state file exists.
 */
export function hasSessionState(): boolean {
  return fs.existsSync(SESSION_STATE_PATH);
}

/**
 * Reads the session state.
 */
export function readSessionState(): SessionState | null {
  return readSecure<SessionState>(SESSION_STATE_PATH);
}

/**
 * Writes the session state securely.
 */
export function writeSessionState(state: SessionState): void {
  ensureConfigDir();
  writeSecure(SESSION_STATE_PATH, state);
}

/**
 * Deletes the session state file.
 */
export function clearSessionState(): void {
  sessionEpoch += 1;
  if (fs.existsSync(SESSION_STATE_PATH)) {
    fs.unlinkSync(SESSION_STATE_PATH);
  }
}

/**
 * Deletes everything stored locally: browser profile (cookies, MSAL caches,
 * every account ever signed in), session state and token cache.
 */
export function clearAllLocalData(): void {
  sessionEpoch += 1;
  fs.rmSync(CONFIG_DIR, { recursive: true, force: true });
}

/**
 * Gets the age of the session state in hours.
 */
export function getSessionAge(): number | null {
  if (!hasSessionState()) {
    return null;
  }

  const stats = fs.statSync(SESSION_STATE_PATH);
  const ageMs = Date.now() - stats.mtimeMs;
  return ageMs / (1000 * 60 * 60);
}

/**
 * Checks if session is likely expired (>12 hours old).
 */
export function isSessionLikelyExpired(): boolean {
  const age = getSessionAge();
  if (age === null) return true;
  return age > SESSION_EXPIRY_HOURS;
}

/**
 * Reads the token cache.
 */
export function readTokenCache(): TokenCache | null {
  return readSecure<TokenCache>(TOKEN_CACHE_PATH);
}

/**
 * Writes the token cache securely.
 */
export function writeTokenCache(cache: TokenCache): void {
  ensureConfigDir();
  writeSecure(TOKEN_CACHE_PATH, cache);
}

/**
 * Clears the token cache.
 */
export function clearTokenCache(): void {
  if (fs.existsSync(TOKEN_CACHE_PATH)) {
    fs.unlinkSync(TOKEN_CACHE_PATH);
  }
}

/**
 * Known Teams origins (commercial and government clouds).
 * Used to find the correct origin in session state.
 */
const TEAMS_ORIGINS = TEAMS_HOSTS.map(host => `https://${host}`);

/**
 * Returns true if localStorage has a SubstrateSearch MSAL token entry.
 */
function hasSubstrateSearchToken(
  localStorage: SessionState['origins'][number]['localStorage'] | undefined
): boolean {
  for (const item of localStorage ?? []) {
    try {
      const entry = JSON.parse(item.value) as { target?: string; secret?: string };
      if (
        entry.target?.includes('SubstrateSearch') &&
        typeof entry.secret === 'string' &&
        entry.secret.startsWith('ey')
      ) {
        return true;
      }
    } catch {
      // ignore non-JSON localStorage values
    }
  }
  return false;
}

/**
 * Latest unexpired JWT exp (seconds) among an origin's localStorage entries.
 * Entries are MSAL JSON blobs; secrets that look like JWTs ('ey' prefix) get
 * their exp claim decoded. Returns null when nothing unexpired exists.
 */
function latestUnexpiredJwtExp(
  localStorage: SessionState['origins'][number]['localStorage'] | undefined
): number | null {
  const nowSec = Date.now() / 1000;
  let best: number | null = null;
  for (const item of localStorage ?? []) {
    try {
      const entry = JSON.parse(item.value) as { secret?: string };
      const secret = entry.secret;
      if (typeof secret !== 'string' || !secret.startsWith('ey')) continue;
      const parts = secret.split('.');
      if (parts.length < 2) continue;
      const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString()) as { exp?: number };
      if (typeof payload.exp !== 'number' || payload.exp <= nowSec) continue;
      if (best === null || payload.exp > best) best = payload.exp;
    } catch {
      // ignore non-JSON or non-JWT values
    }
  }
  return best;
}

/**
 * Gets the Teams origin from session state.
 * Checks multiple known Teams domains to support government clouds.
 *
 * After the New Teams host migration a session can hold TWO Teams origins:
 * a fresh one on teams.cloud.microsoft and a stale one on
 * teams.microsoft.com. Pick the origin whose localStorage contains the
 * latest-expiring UNEXPIRED JWT, so extraction never lands on the stale
 * copy. When nothing is unexpired, fall back to preferring an origin that
 * has a SubstrateSearch token, then known origins in TEAMS_HOSTS order.
 */
export function getTeamsOrigin(state: SessionState): SessionState['origins'][number] | null {
  if (!state.origins) return null;

  let bestOrigin: SessionState['origins'][number] | null = null;
  let bestExp: number | null = null;
  for (const origin of state.origins) {
    if (!TEAMS_HOSTS.some(host => origin.origin.includes(host))) continue;
    const exp = latestUnexpiredJwtExp(origin.localStorage);
    if (exp !== null && (bestExp === null || exp > bestExp)) {
      bestExp = exp;
      bestOrigin = origin;
    }
  }
  if (bestOrigin) return bestOrigin;

  const withSubstrate = state.origins.find(o => hasSubstrateSearchToken(o.localStorage));
  if (withSubstrate) return withSubstrate;

  // Try known Teams origins in priority order
  for (const knownOrigin of TEAMS_ORIGINS) {
    const origin = state.origins.find(o => o.origin === knownOrigin);
    if (origin) return origin;
  }

  // Fallback: any origin on a known Teams host
  return state.origins.find(o => TEAMS_HOSTS.some(host => o.origin.includes(host))) ?? null;
}
