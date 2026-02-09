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
import { fileURLToPath } from 'url';
import { encrypt, decrypt, isEncrypted } from './crypto.js';
import { SESSION_EXPIRY_HOURS } from '../constants.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

/**
 * Gets the user's home directory with fallback.
 * os.homedir() can throw in rare edge cases (missing env vars, broken passwd).
 */
function getHomeDirSafe(): string | null {
  try {
    return os.homedir();
  } catch {
    return null;
  }
}

/**
 * Gets the user-specific config directory for teams-mcp-server.
 * - Windows: %APPDATA%\teams-mcp-server\ (e.g., C:\Users\name\AppData\Roaming\teams-mcp-server\)
 * - macOS/Linux: ~/.teams-mcp-server/
 * - Fallback: ./teams-mcp-server-data/ relative to package (legacy behaviour)
 */
function getConfigDir(): string {
  const homeDir = getHomeDirSafe();
  
  if (process.platform === 'win32') {
    const appData = process.env.APPDATA || (homeDir ? path.join(homeDir, 'AppData', 'Roaming') : null);
    if (appData) {
      return path.join(appData, 'teams-mcp-server');
    }
  } else if (homeDir) {
    return path.join(homeDir, '.teams-mcp-server');
  }
  
  // Fallback to package-relative directory if home directory unavailable
  const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
  return path.join(projectRoot, 'teams-mcp-server-data');
}

export const PROJECT_ROOT = path.resolve(__dirname, '../..');
export const CONFIG_DIR = getConfigDir();
export const USER_DATA_DIR = path.join(CONFIG_DIR, '.user-data');
export const SESSION_STATE_PATH = path.join(CONFIG_DIR, 'session-state.json');
export const TOKEN_CACHE_PATH = path.join(CONFIG_DIR, 'token-cache.json');

// Legacy paths for migration
const LEGACY_SESSION_PATH = path.join(PROJECT_ROOT, 'session-state.json');
const LEGACY_TOKEN_CACHE_PATH = path.join(PROJECT_ROOT, 'token-cache.json');

/** File permission mode: owner read/write only. */
const SECURE_FILE_MODE = 0o600;

/**
 * Ensures the config directory exists with secure permissions.
 */
function ensureConfigDir(): void {
  if (!fs.existsSync(CONFIG_DIR)) {
    fs.mkdirSync(CONFIG_DIR, { recursive: true, mode: 0o700 });
  }
}

/**
 * Migrates a file from legacy location to new config directory.
 * Only migrates if legacy exists and new location doesn't.
 */
function migrateIfNeeded(legacyPath: string, newPath: string): void {
  if (fs.existsSync(legacyPath) && !fs.existsSync(newPath)) {
    ensureConfigDir();
    try {
      fs.copyFileSync(legacyPath, newPath);
      fs.chmodSync(newPath, SECURE_FILE_MODE);
      // Remove legacy file after successful copy
      fs.unlinkSync(legacyPath);
    } catch (error) {
      // Log migration errors for debugging, but continue - will just create new session
      console.error(
        `Failed to migrate ${path.basename(legacyPath)}:`,
        error instanceof Error ? error.message : error
      );
    }
  }
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
 * Ensures the user data directory exists.
 */
export function ensureUserDataDir(): void {
  ensureConfigDir();
  if (!fs.existsSync(USER_DATA_DIR)) {
    fs.mkdirSync(USER_DATA_DIR, { recursive: true, mode: 0o700 });
  }
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
 * Reads data securely, handling both encrypted and legacy plaintext.
 */
function readSecure<T>(filePath: string): T | null {
  if (!fs.existsSync(filePath)) {
    return null;
  }

  try {
    const content = fs.readFileSync(filePath, 'utf8');
    const parsed = JSON.parse(content);

    // Check if this is encrypted data
    if (isEncrypted(parsed)) {
      const decrypted = decrypt(parsed);
      return JSON.parse(decrypted) as T;
    }

    // Legacy plaintext - migrate to encrypted
    writeSecure(filePath, parsed);
    return parsed as T;

  } catch (error) {
    // If decryption fails (different machine, corrupted), return null
    console.error(`Failed to read ${filePath}:`, error instanceof Error ? error.message : error);
    return null;
  }
}

/**
 * Checks if session state file exists.
 */
export function hasSessionState(): boolean {
  migrateIfNeeded(LEGACY_SESSION_PATH, SESSION_STATE_PATH);
  return fs.existsSync(SESSION_STATE_PATH);
}

/**
 * Reads the session state.
 */
export function readSessionState(): SessionState | null {
  migrateIfNeeded(LEGACY_SESSION_PATH, SESSION_STATE_PATH);
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
  migrateIfNeeded(LEGACY_TOKEN_CACHE_PATH, TOKEN_CACHE_PATH);
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
  'https://teams.microsoft.com',   // Commercial
  'https://teams.microsoft.us',    // GCC-High
  'https://dod.teams.microsoft.us', // DoD
  'https://teams.cloud.microsoft', // New Teams URL
];

/**
 * Gets the Teams origin from session state.
 * Checks multiple known Teams domains to support government clouds.
 */
export function getTeamsOrigin(state: SessionState): SessionState['origins'][number] | null {
  if (!state.origins) return null;
  
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
