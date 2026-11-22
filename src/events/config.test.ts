/**
 * Unit tests for the MCP Events opt-in and interval clamping.
 */

import { describe, it, expect } from 'vitest';
import {
  clampEventsIntervalMs,
  parseEventsEnabled,
  resolveEventsConfig,
  EVENTS_ENABLED_ENV,
  EVENTS_INTERVAL_ENV,
} from './config.js';
import {
  DEFAULT_EVENTS_INTERVAL_MS,
  MAX_EVENTS_INTERVAL_MS,
  MIN_EVENTS_INTERVAL_MS,
} from '../constants.js';

describe('parseEventsEnabled', () => {
  it('accepts the documented truthy values, case-insensitively', () => {
    for (const value of ['1', 'true', 'TRUE', 'yes', 'on', ' on ']) {
      expect(parseEventsEnabled(value)).toBe(true);
    }
  });

  it('treats absent, empty and other values as disabled', () => {
    for (const value of [undefined, '', '0', 'false', 'no', 'off', 'maybe']) {
      expect(parseEventsEnabled(value)).toBe(false);
    }
  });
});

describe('clampEventsIntervalMs', () => {
  it('defaults an absent or invalid interval', () => {
    expect(clampEventsIntervalMs(undefined)).toBe(DEFAULT_EVENTS_INTERVAL_MS);
    expect(clampEventsIntervalMs(Number.NaN)).toBe(DEFAULT_EVENTS_INTERVAL_MS);
    expect(clampEventsIntervalMs(0)).toBe(DEFAULT_EVENTS_INTERVAL_MS);
    expect(clampEventsIntervalMs(-5)).toBe(DEFAULT_EVENTS_INTERVAL_MS);
  });

  it('enforces the traffic floor and the maximum', () => {
    expect(clampEventsIntervalMs(1000)).toBe(MIN_EVENTS_INTERVAL_MS);
    expect(clampEventsIntervalMs(MAX_EVENTS_INTERVAL_MS + 1)).toBe(MAX_EVENTS_INTERVAL_MS);
  });

  it('keeps an interval inside the supported range', () => {
    expect(clampEventsIntervalMs(45_000)).toBe(45_000);
  });
});

describe('resolveEventsConfig', () => {
  it('is disabled by default', () => {
    const config = resolveEventsConfig({});
    expect(config.enabled).toBe(false);
    expect(config.intervalMs).toBe(DEFAULT_EVENTS_INTERVAL_MS);
  });

  it('reads both environment variables', () => {
    const config = resolveEventsConfig({
      [EVENTS_ENABLED_ENV]: 'true',
      [EVENTS_INTERVAL_ENV]: '30000',
    });
    expect(config.enabled).toBe(true);
    expect(config.intervalMs).toBe(30_000);
  });

  it('clamps a dangerously small configured interval', () => {
    const config = resolveEventsConfig({ [EVENTS_INTERVAL_ENV]: '50' });
    expect(config.intervalMs).toBe(MIN_EVENTS_INTERVAL_MS);
  });
});
