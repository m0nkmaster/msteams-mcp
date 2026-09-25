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
import { SESSION_EXPIRY_HOURS } from '../constants.js';
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
  if (fs.existsSync(SESSION_STATE_PATH)) {
    fs.unlinkSync(SESSION_STATE_PATH);
  }
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
const TEAMS_ORIGINS = [
  'https://teams.cloud.microsoft', // New Teams URL (often holds SubstrateSearch tokens)
  'https://teams.microsoft.com',   // Commercial legacy
  'https://teams.microsoft.us',    // GCC-High
  'https://dod.teams.microsoft.us', // DoD
];

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
 * Gets the Teams origin from session state.
 * Checks multiple known Teams domains to support government clouds.
 *
 * Prefers an origin that has a SubstrateSearch token so Outlook/email and
 * Substrate search keep working when both teams.cloud.microsoft and
 * teams.microsoft.com are present in the saved session (common after the
 * New Teams host migration).
 */
export function getTeamsOrigin(state: SessionState): SessionState['origins'][number] | null {
  if (!state.origins) return null;

  const withSubstrate = state.origins.find(o => hasSubstrateSearchToken(o.localStorage));
  if (withSubstrate) return withSubstrate;

  // Try known Teams origins in priority order
  for (const knownOrigin of TEAMS_ORIGINS) {
    const origin = state.origins.find(o => o.origin === knownOrigin);
    if (origin) return origin;
  }

  // Fallback: find any origin containing 'teams.microsoft'
  return state.origins.find(o =>
    o.origin.includes('teams.microsoft') || o.origin.includes('teams.cloud')
  ) ?? null;
}
