/**
 * Lightweight logger for consistent log output.
 * 
 * All production logging goes through this module so that:
 * - Output is consistently prefixed with `[module]` context
 * - Log levels can be controlled via LOG_LEVEL environment variable
 * - Future structured logging can be added in one place
 * 
 * Levels: error | warn | info (default) | debug
 * Set LOG_LEVEL=debug for verbose output during development.
 */

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

/** Log an error message. Always shown unless level is somehow above error. */
export function error(context: string, message: string): void {
  if (shouldLog('error')) {
    console.error(`[${context}] ${message}`);
  }
}

/** Log a warning message. */
export function warn(context: string, message: string): void {
  if (shouldLog('warn')) {
    console.warn(`[${context}] ${message}`);
  }
}

/** Log an informational message. */
export function info(context: string, message: string): void {
  if (shouldLog('info')) {
    console.log(`[${context}] ${message}`);
  }
}

/** Log a debug message. Only shown when LOG_LEVEL=debug. */
export function debug(context: string, message: string): void {
  if (shouldLog('debug')) {
    console.debug(`[${context}] ${message}`);
  }
}
