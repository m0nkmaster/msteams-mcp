/**
 * Promise deadline helpers.
 *
 * Used to bound browser and network operations that can wedge silently
 * (unresponsive renderer, dead socket) so a single stuck call can never
 * block a tool call or the serialized token-refresh queue forever.
 */

/** Error thrown when a promise does not settle within its deadline. */
export class TimeoutError extends Error {
  constructor(label: string, timeoutMs: number) {
    super(`${label} timed out after ${timeoutMs}ms`);
    this.name = 'TimeoutError';
  }
}

/**
 * Races `promise` against a deadline. On timeout the returned promise rejects
 * with TimeoutError; the underlying promise keeps running in the background
 * (Promise.race subscribes to it, so a late rejection is not unhandled).
 * Callers that need the underlying work stopped must clean it up separately.
 */
export function withTimeout<T>(promise: Promise<T>, timeoutMs: number, label: string): Promise<T> {
  let timer: NodeJS.Timeout | undefined;
  const deadline = new Promise<never>((_, reject) => {
    timer = setTimeout(() => reject(new TimeoutError(label, timeoutMs)), timeoutMs);
    timer.unref?.();
  });
  return Promise.race([promise, deadline]).finally(() => clearTimeout(timer));
}
