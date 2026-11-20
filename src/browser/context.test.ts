import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('playwright', () => ({
  chromium: { launchPersistentContext: vi.fn() },
}));

vi.mock('../auth/session-store.js', () => ({
  ensureConfigDir: vi.fn(),
  CONFIG_DIR: '/fake/config',
  writeSessionState: vi.fn(),
}));

vi.mock('../utils/auth-guards.js', () => ({ clearRegionCache: vi.fn() }));
vi.mock('./force-kill.js', () => ({ killBrowserProfileProcesses: vi.fn() }));
vi.mock('../utils/logger.js', () => ({
  info: vi.fn(),
  warn: vi.fn(),
  error: vi.fn(),
  debug: vi.fn(),
}));

import { chromium } from 'playwright';
import { createBrowserContext, closeBrowser, saveSessionState } from './context.js';
import { writeSessionState } from '../auth/session-store.js';
import { killBrowserProfileProcesses } from './force-kill.js';
import * as log from '../utils/logger.js';

function mockLaunchedContext() {
  const context = {
    pages: () => [{}],
    newPage: vi.fn(),
    storageState: vi.fn().mockResolvedValue({ cookies: [], origins: [] }),
    close: vi.fn().mockResolvedValue(undefined),
  };
  vi.mocked(chromium.launchPersistentContext).mockResolvedValue(context as never);
  return context;
}

describe('createBrowserContext', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('passes the launch timeout through to Playwright when provided', async () => {
    mockLaunchedContext();

    await createBrowserContext({ headless: true, launchTimeoutMs: 30_000 });

    const options = vi.mocked(chromium.launchPersistentContext).mock.calls[0][1];
    expect(options).toMatchObject({ headless: true, timeout: 30_000 });
  });

  it('omits the launch timeout by default', async () => {
    mockLaunchedContext();

    await createBrowserContext({ headless: true });

    const options = vi.mocked(chromium.launchPersistentContext).mock.calls[0][1];
    expect(options?.timeout).toBeUndefined();
  });

  it('warns clearly when the profile is held by another process', async () => {
    vi.mocked(chromium.launchPersistentContext).mockRejectedValue(
      new Error('browserType.launchPersistentContext: Target page, context or browser has been closed'),
    );

    await expect(createBrowserContext({ headless: true })).rejects.toThrow('Could not launch');
    expect(log.warn).toHaveBeenCalledWith(
      'browser',
      expect.stringContaining('in use by another process'),
    );
  });
});

describe('saveSessionState', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it('bounds a wedged storageState() and does not write partial state', async () => {
    vi.useFakeTimers();
    const context = {
      storageState: vi.fn(() => new Promise(() => {})),
    };

    const pending = saveSessionState(context as never);
    const assertion = expect(pending).rejects.toThrow('timed out');
    await vi.advanceTimersByTimeAsync(15_000);
    await assertion;

    expect(writeSessionState).not.toHaveBeenCalled();
  });
});

describe('closeBrowser', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it('force-kills profile processes when close() wedges', async () => {
    vi.useFakeTimers();
    const manager = {
      context: { close: vi.fn(() => new Promise(() => {})) },
      page: {},
    } as never;

    const pending = closeBrowser(manager, false);
    const assertion = expect(pending).rejects.toThrow('timed out');
    await vi.advanceTimersByTimeAsync(15_000);
    await assertion;

    expect(killBrowserProfileProcesses).toHaveBeenCalledWith(expect.stringContaining('browser-profile'));
  });

  it('propagates a normal close error without killing processes', async () => {
    const manager = {
      context: { close: vi.fn().mockRejectedValue(new Error('Target closed')) },
      page: {},
    } as never;

    await expect(closeBrowser(manager, false)).rejects.toThrow('Target closed');
    expect(killBrowserProfileProcesses).not.toHaveBeenCalled();
  });
});
