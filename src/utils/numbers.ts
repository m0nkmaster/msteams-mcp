/**
 * Small numeric helpers shared by option parsing.
 */

/**
 * Clamps an integer option into `[min, max]`, falling back to `fallback` when
 * the value is absent or not a finite number. Used for scan sizes, poll
 * intervals and heartbeat intervals so every option behaves the same way.
 */
export function clampInt(
  value: number | undefined,
  fallback: number,
  min: number,
  max: number
): number {
  if (value === undefined || !Number.isFinite(value)) return fallback;
  return Math.min(Math.max(Math.floor(value), min), max);
}
