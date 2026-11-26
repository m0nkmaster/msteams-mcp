import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { ErrorCode, createError } from '../types/errors.js';

vi.mock('./token-extractor.js', () => ({
  extractSubstrateToken: vi.fn(),
  extractMessageAuth: vi.fn(),
  extractSkypeSpacesToken: vi.fn(),
  extractCsaToken: vi.fn(),
  getValidAssignmentsToken: vi.fn(),
  getValidGraphToken: vi.fn(),
}));

vi.mock('./session-store.js', () => ({
  clearTokenCache: vi.fn(),
}));

vi.mock('../browser/context.js', () => ({ createBrowserContext: vi.fn(), closeBrowser: vi.fn() }));
vi.mock('../browser/auth.js', () => ({ ensureAuthenticated: vi.fn() }));

vi.mock('./token-refresh-http.js', () => ({
  refreshTokensViaHttp: vi.fn(),
}));

import { refreshTokensViaBrowser, refreshAssignmentsToken, refreshCoreTokensViaHttp } from './token-refresh.js';
import {
  extractSubstrateToken,
  extractMessageAuth,
  extractSkypeSpacesToken,
  extractCsaToken,
} from './token-extractor.js';
import { refreshTokensViaHttp } from './token-refresh-http.js';
import type { Result } from '../types/result.js';
import { createBrowserContext, closeBrowser } from '../browser/context.js';
import { ensureAuthenticated } from '../browser/auth.js';
import { CORE_REFRESH_TIMEOUT_MS } from '../constants.js';

const VALID_SUBSTRATE = { token: 'substrate-jwt', expiry: new Date(Date.now() + 3600_000) };
const VALID_MESSAGE_AUTH = { skypeToken: 'skype', authToken: 'auth', userMri: '8:orgid:u' };

/** Marks every core credential extractor as valid. */
function completeSession(): void {
  vi.mocked(extractSubstrateToken).mockReturnValue(VALID_SUBSTRATE);
  vi.mocked(extractMessageAuth).mockReturnValue(VALID_MESSAGE_AUTH);
  vi.mocked(extractSkypeSpacesToken).mockReturnValue('spaces');
  vi.mocked(extractCsaToken).mockReturnValue('csa');
}

describe('refreshTokensViaBrowser', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it('uses HTTP refresh after the cached access token has expired', async () => {
    completeSession();
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({ ok: true, value: undefined });

    const result = await refreshTokensViaBrowser();

    expect(refreshTokensViaHttp).toHaveBeenCalledOnce();
    expect(createBrowserContext).not.toHaveBeenCalled();
    expect(result.ok).toBe(true);
  });

  it('returns the refresh error when HTTP refresh fails and no browser is mocked', async () => {
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({
      ok: false,
      error: createError(ErrorCode.AUTH_REQUIRED, 'no refresh token'),
    });
    vi.mocked(createBrowserContext).mockRejectedValue(new Error('no browser'));

    const result = await refreshTokensViaBrowser();

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.message).toContain('Token refresh via browser failed');
    }
  });

  it('abandons a never-settling refresh at the deadline and lets the next call retry', async () => {
    vi.useFakeTimers();

    // First attempt hangs forever inside the HTTP refresh.
    vi.mocked(refreshTokensViaHttp).mockImplementation(() => new Promise(() => {}));

    const hung = refreshTokensViaBrowser();
    await vi.advanceTimersByTimeAsync(CORE_REFRESH_TIMEOUT_MS);

    const hungResult = await hung;
    expect(hungResult.ok).toBe(false);
    if (!hungResult.ok) {
      expect(hungResult.error.code).toBe(ErrorCode.AUTH_REQUIRED);
      expect(hungResult.error.message).toContain('timed out after 90s');
      expect(hungResult.error.message).toContain('teams_login');
    }

    // The queue advanced: a later call starts a fresh attempt and succeeds.
    completeSession();
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({ ok: true, value: undefined });

    const retry = await refreshTokensViaBrowser();
    expect(retry.ok).toBe(true);
    expect(vi.mocked(refreshTokensViaHttp)).toHaveBeenCalledTimes(2);
  });

  it('queues refreshCoreTokensViaHttp behind an in-flight core refresh', async () => {
    completeSession();
    let resolveFirst!: (value: Result<void>) => void;
    const firstCall = new Promise<Result<void>>(resolve => { resolveFirst = resolve; });
    vi.mocked(refreshTokensViaHttp)
      .mockImplementationOnce(() => firstCall)
      .mockResolvedValue({ ok: true, value: undefined });

    const browserRefresh = refreshTokensViaBrowser();
    let httpRefresh: Promise<Result<void>> | undefined;
    try {
      httpRefresh = refreshCoreTokensViaHttp();

      // The queued HTTP refresh must not reach refreshTokensViaHttp while the
      // first attempt is still in flight.
      await new Promise(resolve => setTimeout(resolve, 10));
      expect(vi.mocked(refreshTokensViaHttp)).toHaveBeenCalledTimes(1);

      resolveFirst({ ok: true, value: undefined });
      const [browserResult, httpResult] = await Promise.all([browserRefresh, httpRefresh]);
      expect(browserResult.ok).toBe(true);
      expect(httpResult.ok).toBe(true);
      expect(vi.mocked(refreshTokensViaHttp)).toHaveBeenCalledTimes(2);
    } finally {
      // Never leave the shared refresh queue wedged on the deferred call.
      resolveFirst({ ok: true, value: undefined });
      await Promise.allSettled([browserRefresh, httpRefresh]);
    }
  });

  it('gives on-demand refreshes the same deadline', async () => {
    vi.useFakeTimers();
    vi.mocked(refreshTokensViaHttp).mockImplementation(() => new Promise(() => {}));

    const pending = refreshAssignmentsToken();
    await vi.advanceTimersByTimeAsync(CORE_REFRESH_TIMEOUT_MS);

    const result = await pending;
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe(ErrorCode.TIMEOUT);
      expect(result.error.message).toContain('timed out after 90s');
    }
  });

  it('tops up missing messaging cookies after a browser refresh that saved no session cookies', async () => {
    // HTTP refresh fails outright → browser fallback runs and "succeeds" but
    // leaves the messaging cookies unminted (the reproduced auth loop).
    // The top-up exchange (second HTTP call) is what restores them.
    let httpCalls = 0;
    let toppedUp = false;
    vi.mocked(refreshTokensViaHttp).mockImplementation(async () => {
      httpCalls++;
      if (httpCalls === 1) {
        return { ok: false as const, error: createError(ErrorCode.AUTH_REQUIRED, 'no refresh token') };
      }
      toppedUp = true;
      return { ok: true as const, value: undefined };
    });
    vi.mocked(createBrowserContext).mockResolvedValue({ page: {}, context: {} } as never);
    vi.mocked(ensureAuthenticated).mockResolvedValue(undefined);
    vi.mocked(closeBrowser).mockResolvedValue(undefined as never);

    // Before the top-up the messaging cookies are absent; the top-up restores them.
    vi.mocked(extractSubstrateToken).mockReturnValue(VALID_SUBSTRATE);
    vi.mocked(extractSkypeSpacesToken).mockReturnValue('spaces');
    vi.mocked(extractCsaToken).mockReturnValue('csa');
    vi.mocked(extractMessageAuth).mockImplementation(() => (toppedUp ? VALID_MESSAGE_AUTH : null));

    const result = await refreshTokensViaBrowser();

    expect(result.ok).toBe(true);
    expect(createBrowserContext).toHaveBeenCalledWith(
      expect.objectContaining({ headless: true, launchTimeoutMs: 30_000 }),
    );
    // One failed HTTP attempt + one top-up exchange.
    expect(vi.mocked(refreshTokensViaHttp)).toHaveBeenCalledTimes(2);
  });

  it('returns AUTH_REQUIRED when the top-up cannot mint messaging cookies', async () => {
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({
      ok: false,
      error: createError(ErrorCode.AUTH_REQUIRED, 'no refresh token'),
    });
    vi.mocked(createBrowserContext).mockResolvedValue({ page: {}, context: {} } as never);
    vi.mocked(ensureAuthenticated).mockResolvedValue(undefined);
    vi.mocked(closeBrowser).mockResolvedValue(undefined as never);

    vi.mocked(extractSubstrateToken).mockReturnValue(VALID_SUBSTRATE);
    vi.mocked(extractSkypeSpacesToken).mockReturnValue('spaces');
    vi.mocked(extractCsaToken).mockReturnValue('csa');
    vi.mocked(extractMessageAuth).mockReturnValue(null);

    const result = await refreshTokensViaBrowser();

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe(ErrorCode.AUTH_REQUIRED);
      expect(result.error.message).toContain('teams_login');
    }
  });
});
