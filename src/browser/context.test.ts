/**
 * Tests for browser selection.
 */

import { describe, it, expect, afterEach, vi } from 'vitest';
import { getBrowserChoice } from './context.js';

/** Runs fn with process.platform stubbed, then restores the real value. */
function withPlatform<T>(platform: NodeJS.Platform, fn: () => T): T {
  const original = Object.getOwnPropertyDescriptor(process, 'platform')!;
  Object.defineProperty(process, 'platform', { value: platform, configurable: true });
  try {
    return fn();
  } finally {
    Object.defineProperty(process, 'platform', original);
  }
}

describe('getBrowserChoice', () => {
  afterEach(() => {
    vi.unstubAllEnvs();
  });

  it('defaults to Edge on Windows and Chrome elsewhere', () => {
    vi.stubEnv('TEAMS_BROWSER', '');
    expect(withPlatform('win32', getBrowserChoice)).toBe('msedge');
    expect(withPlatform('darwin', getBrowserChoice)).toBe('chrome');
    expect(withPlatform('linux', getBrowserChoice)).toBe('chrome');
  });

  it('honours TEAMS_BROWSER on every platform', () => {
    vi.stubEnv('TEAMS_BROWSER', ' Firefox ');
    expect(withPlatform('linux', getBrowserChoice)).toBe('firefox');
    vi.stubEnv('TEAMS_BROWSER', 'msedge');
    expect(withPlatform('darwin', getBrowserChoice)).toBe('msedge');
    vi.stubEnv('TEAMS_BROWSER', 'chrome');
    expect(withPlatform('win32', getBrowserChoice)).toBe('chrome');
  });

  it('falls back to the platform default for an unknown value', () => {
    vi.stubEnv('TEAMS_BROWSER', 'safari');
    expect(withPlatform('linux', getBrowserChoice)).toBe('chrome');
  });
});
