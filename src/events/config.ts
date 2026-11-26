/**
 * MCP Events configuration.
 *
 * Event delivery is opt-in because it makes this server poll Teams on a timer
 * from the user's machine. The interval has a floor so a misconfigured client
 * or environment cannot turn the server into a traffic generator.
 */

import {
  DEFAULT_EVENTS_INTERVAL_MS,
  MAX_EVENTS_INTERVAL_MS,
  MIN_EVENTS_INTERVAL_MS,
} from '../constants.js';
import { clampInt } from '../utils/numbers.js';

/** Set to a truthy value (`1`, `true`, `yes`, `on`) to enable MCP Events. */
export const EVENTS_ENABLED_ENV = 'TEAMS_EVENTS';

/** Upstream poll interval in milliseconds (clamped to the supported range). */
export const EVENTS_INTERVAL_ENV = 'TEAMS_EVENTS_INTERVAL_MS';

/** Settings resolved from the environment for the events subsystem. */
export interface EventsConfig {
  enabled: boolean;
  /** Interval between upstream Teams polls. */
  intervalMs: number;
}

/** Parses a boolean-ish environment value. Absent or unknown means disabled. */
export function parseEventsEnabled(value: string | undefined): boolean {
  return /^(1|true|yes|on)$/i.test((value ?? '').trim());
}

/** Clamps a requested interval into the supported range. */
export function clampEventsIntervalMs(value: number | undefined): number {
  if (value !== undefined && (!Number.isFinite(value) || value <= 0)) {
    return DEFAULT_EVENTS_INTERVAL_MS;
  }
  return clampInt(value, DEFAULT_EVENTS_INTERVAL_MS, MIN_EVENTS_INTERVAL_MS, MAX_EVENTS_INTERVAL_MS);
}

/**
 * Resolves the events configuration from an environment object.
 * Pure so the opt-in rules and clamping are unit-testable.
 */
export function resolveEventsConfig(env: NodeJS.ProcessEnv = process.env): EventsConfig {
  const rawInterval = env[EVENTS_INTERVAL_ENV];
  const parsedInterval = rawInterval === undefined ? undefined : Number(rawInterval);

  return {
    enabled: parseEventsEnabled(env[EVENTS_ENABLED_ENV]),
    intervalMs: clampEventsIntervalMs(parsedInterval),
  };
}
