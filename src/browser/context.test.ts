/**
 * Tests for browser channel selection.
 *
 * The server launches into its own profile dir, so the user's default browser
 * is irrelevant - what matters is that every platform has an ordered list of
 * channels to try and ends with a fallback that needs no system browser.
 */

import { describe, it, expect, afterEach, vi } from 'vitest';
import { browserChannels } from './context.js';

/** Runs fn with process.platform stubbed, then restores the real value. */
function withPlatform<T>(platform: string, fn: () => T): T {
  const original = Object.getOwnPropertyDescriptor(process, 'platform')!;
  Object.defineProperty(process, 'platform', { value: platform, configurable: true });
  try {
    return fn();
  } finally {
    Object.defineProperty(process, 'platform', original);
  }
}

describe('browserChannels', () => {
  afterEach(() => {
    vi.unstubAllEnvs();
  });

  it('prefers Edge on Windows, Chrome elsewhere', () => {
    vi.stubEnv('TEAMS_BROWSER_CHANNEL', '');

    expect(withPlatform('win32', browserChannels)[0]).toBe('msedge');
    expect(withPlatform('darwin', browserChannels)[0]).toBe('chrome');
    expect(withPlatform('linux', browserChannels)[0]).toBe('chrome');
  });

  it('offers both system channels on every platform', () => {
    vi.stubEnv('TEAMS_BROWSER_CHANNEL', '');

    for (const platform of ['win32', 'darwin', 'linux']) {
      const channels = withPlatform(platform, browserChannels);
      expect(channels).toContain('chrome');
      expect(channels).toContain('msedge');
    }
  });

  it('ends with bundled Chromium so a machine with no system browser can recover', () => {
    vi.stubEnv('TEAMS_BROWSER_CHANNEL', '');

    for (const platform of ['win32', 'darwin', 'linux']) {
      expect(withPlatform(platform, browserChannels).at(-1)).toBeUndefined();
    }
  });

  it('honours TEAMS_BROWSER_CHANNEL and disables the fallback', () => {
    vi.stubEnv('TEAMS_BROWSER_CHANNEL', 'msedge');
    expect(withPlatform('darwin', browserChannels)).toEqual(['msedge']);

    vi.stubEnv('TEAMS_BROWSER_CHANNEL', 'chrome');
    expect(withPlatform('win32', browserChannels)).toEqual(['chrome']);
  });

  it('ignores an unrecognised TEAMS_BROWSER_CHANNEL', () => {
    vi.stubEnv('TEAMS_BROWSER_CHANNEL', 'safari');
    expect(withPlatform('darwin', browserChannels)).toEqual(['chrome', 'msedge', undefined]);
  });
});
