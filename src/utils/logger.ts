/**
 * Lightweight logger for consistent log output.
 * 
 * All production logging goes through this module so that:
 * - Everything goes to stderr: stdout is the MCP stdio protocol channel, and
 *   the CLI's data output
 * - Output is consistently prefixed with `[module]` context
 * - Log levels can be controlled via LOG_LEVEL environment variable
 * - Every emitted line can also be appended to a file (see setLogFilePath)
 *   for environments that do not retain the server's stderr
 * 
 * Levels: error | warn | info (default) | debug
 * Set LOG_LEVEL=debug for verbose output during development.
 */

import * as fs from 'fs';
import * as path from 'path';

type LogLevel = 'error' | 'warn' | 'info' | 'debug';

const LEVEL_PRIORITY: Record<LogLevel, number> = {
  error: 0,
  warn: 1,
  info: 2,
  debug: 3,
};

function getConfiguredLevel(): LogLevel {
  const env = process.env.LOG_LEVEL?.toLowerCase();
  if (env && env in LEVEL_PRIORITY) return env as LogLevel;
  return 'info';
}

let currentLevel = getConfiguredLevel();

function shouldLog(level: LogLevel): boolean {
  return LEVEL_PRIORITY[level] <= LEVEL_PRIORITY[currentLevel];
}

/** Set the log level at runtime (useful for testing). */
export function setLogLevel(level: LogLevel): void {
  currentLevel = level;
}

// ─────────────────────────────────────────────────────────────────────────────
// File logging
// ─────────────────────────────────────────────────────────────────────────────

/** Rotate the log file once it grows past this size (bytes). */
const LOG_FILE_MAX_BYTES = 1024 * 1024;

/** Active log file path, or null when file logging is disabled. */
let logFilePath: string | null = null;

/**
 * Enables appending every emitted line to `filePath`, ISO-timestamped.
 * When the file exceeds 1 MB it is rotated to `<filePath>.1` (one backup
 * kept). Pass null to disable. File write failures are swallowed — logging
 * must never break the server.
 */
export function setLogFilePath(filePath: string | null): void {
  logFilePath = filePath;
}

function appendToFile(line: string): void {
  if (!logFilePath) return;
  try {
    fs.mkdirSync(path.dirname(logFilePath), { recursive: true });
    try {
      if (fs.statSync(logFilePath).size > LOG_FILE_MAX_BYTES) {
        fs.renameSync(logFilePath, `${logFilePath}.1`);
      }
    } catch {
      // Missing file or failed rotation — appendFileSync recreates as needed.
    }
    fs.appendFileSync(logFilePath, `${new Date().toISOString()} ${line}\n`, 'utf8');
  } catch {
    // File logging is best-effort diagnostics only.
  }
}

/** Log an error message. Always shown unless level is somehow above error. */
export function error(context: string, message: string): void {
  if (shouldLog('error')) {
    const line = `[${context}] ${message}`;
    console.error(line);
    appendToFile(line);
  }
}

/** Log a warning message. */
export function warn(context: string, message: string): void {
  if (shouldLog('warn')) {
    const line = `[${context}] ${message}`;
    console.warn(line);
    appendToFile(line);
  }
}

/** Log an informational message. */
export function info(context: string, message: string): void {
  if (shouldLog('info')) {
    const line = `[${context}] ${message}`;
    console.error(line);
    appendToFile(line);
  }
}

/** Log a debug message. Only shown when LOG_LEVEL=debug. */
export function debug(context: string, message: string): void {
  if (shouldLog('debug')) {
    const line = `[${context}] ${message}`;
    console.error(line);
    appendToFile(line);
  }
}
