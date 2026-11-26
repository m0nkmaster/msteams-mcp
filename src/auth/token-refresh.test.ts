import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { ErrorCode, createError } from '../types/errors.js';

vi.mock('./token-extractor.js', () => ({
  extractSubstrateToken: vi.fn(),
  extractMessageAuth: vi.fn(),
  extractSkypeSpacesToken: vi.fn(),
  extractCsaToken: vi.fn(),
  getMessageAuthStatus: vi.fn(),
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

import { awaitRefreshQueue, refreshTokensViaBrowser, refreshAssignmentsToken, refreshCoreTokensViaHttp, refreshGraphToken, setBrowserLoginInProgress } from './token-refresh.js';
import {
  extractSubstrateToken,
  extractMessageAuth,
  extractSkypeSpacesToken,
  extractCsaToken,
  getMessageAuthStatus,
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
  vi.mocked(getMessageAuthStatus).mockReturnValue({ hasToken: true, minutesRemaining: 600 });
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

  it('skips the HTTP exchange when an in-flight refresh already left the session ready', async () => {
    // Session starts incomplete and only becomes ready when the first
    // refresh resolves — an early coreSessionReady() check outside the
    // serialized job would fire a second exchange and fail this test.
    vi.mocked(extractMessageAuth).mockReturnValue(null);
    vi.mocked(getMessageAuthStatus).mockReturnValue({ hasToken: false });
    let resolveFirst!: (value: Result<void>) => void;
    const firstCall = new Promise<Result<void>>(resolve => { resolveFirst = resolve; });
    vi.mocked(refreshTokensViaHttp)
      .mockImplementationOnce(async () => {
        const result = await firstCall;
        completeSession();
        return result;
      })
      .mockResolvedValue({ ok: true, value: undefined });

    const browserRefresh = refreshTokensViaBrowser();
    let httpRefresh: Promise<Result<void>> | undefined;
    let httpSettled = false;
    try {
      httpRefresh = refreshCoreTokensViaHttp().then(result => {
        httpSettled = true;
        return result;
      });

      // The queued HTTP refresh must not reach refreshTokensViaHttp while the
      // first attempt is still in flight, and it must still be pending then.
      await new Promise(resolve => setTimeout(resolve, 10));
      expect(httpSettled).toBe(false);
      expect(vi.mocked(refreshTokensViaHttp)).toHaveBeenCalledTimes(1);

      resolveFirst({ ok: true, value: undefined });
      const [browserResult, httpResult] = await Promise.all([browserRefresh, httpRefresh]);
      expect(browserResult.ok).toBe(true);
      expect(httpResult.ok).toBe(true);
      // The queued job re-checks the session first: the in-flight refresh
      // already minted every credential, so no second OAuth exchange runs.
      expect(vi.mocked(refreshTokensViaHttp)).toHaveBeenCalledTimes(1);
    } finally {
      // Never leave the shared refresh queue wedged on the deferred call.
      resolveFirst({ ok: true, value: undefined });
      await Promise.allSettled([browserRefresh, httpRefresh]);
    }
  });

  it('awaitRefreshQueue waits for wedged-browser cleanup after a refresh timeout', async () => {
    vi.useFakeTimers();
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({
      ok: false,
      error: createError(ErrorCode.AUTH_REQUIRED, 'no refresh token'),
    });
    vi.mocked(createBrowserContext).mockResolvedValue({ page: {}, context: {} } as never);
    // The browser fallback hangs, so the deadline abandons it and the
    // catch-path cleanup (closeBrowser -> profile kill) is still running
    // when the queue tail settles. awaitRefreshQueue must wait for THAT.
    vi.mocked(ensureAuthenticated).mockImplementationOnce(() => new Promise(() => {}));
    let resolveClose!: () => void;
    const closeGate = new Promise<void>(resolve => { resolveClose = resolve; });
    vi.mocked(closeBrowser).mockImplementation(() => closeGate as never);

    const browserRefresh = refreshTokensViaBrowser();
    let queueSettled = false;
    const queueWait = awaitRefreshQueue().then(() => { queueSettled = true; });
    try {
      await vi.advanceTimersByTimeAsync(CORE_REFRESH_TIMEOUT_MS);
      await vi.advanceTimersByTimeAsync(0);

      // Queue tail settled at the deadline, but cleanup still holds the
      // profile — the barrier must not have resolved yet.
      expect(queueSettled).toBe(false);

      resolveClose();
      await Promise.allSettled([browserRefresh, queueWait]);
      expect(queueSettled).toBe(true);
    } finally {
      resolveClose();
      await Promise.allSettled([browserRefresh, queueWait]);
      vi.mocked(closeBrowser).mockResolvedValue(undefined as never);
      vi.useRealTimers();
    }
  });

  it('skips the browser fallback while a teams_login browser session is in progress', async () => {
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({
      ok: false,
      error: createError(ErrorCode.AUTH_REQUIRED, 'no refresh token'),
    });
    setBrowserLoginInProgress(true);
    try {
      const result = await refreshTokensViaBrowser();

      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.error.code).toBe(ErrorCode.AUTH_REQUIRED);
      }
      // A login browser already owns the profile: launching a second one
      // would fail on the Chromium lock anyway — fail fast instead.
      expect(createBrowserContext).not.toHaveBeenCalled();
    } finally {
      setBrowserLoginInProgress(false);
    }
  });

  it('fails fast at entry while a teams_login browser session is in progress', async () => {
    setBrowserLoginInProgress(true);
    try {
      const result = await refreshTokensViaBrowser();

      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.error.code).toBe(ErrorCode.AUTH_REQUIRED);
      }
      // The gate is at entry, not just the browser fallback: the HTTP leg
      // writes session state and must not run while a sign-in owns the profile.
      expect(refreshTokensViaHttp).not.toHaveBeenCalled();
      expect(createBrowserContext).not.toHaveBeenCalled();
    } finally {
      setBrowserLoginInProgress(false);
    }
  });

  it('skips the browser fallback when a sign-in starts while the HTTP leg is in flight', async () => {
    let resolveHttp!: (value: Result<void>) => void;
    vi.mocked(refreshTokensViaHttp).mockImplementation(
      () => new Promise<Result<void>>(resolve => { resolveHttp = resolve; })
    );
    const pending = refreshTokensViaBrowser();
    await vi.waitFor(() => expect(refreshTokensViaHttp).toHaveBeenCalledTimes(1));

    // The refresh passed its entry gate before the sign-in raised the flag;
    // the fallback must re-check before touching the browser profile.
    setBrowserLoginInProgress(true);
    try {
      resolveHttp({ ok: false, error: createError(ErrorCode.AUTH_REQUIRED, 'refresh token rejected') });
      const result = await pending;

      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.error.code).toBe(ErrorCode.AUTH_REQUIRED);
      }
      expect(createBrowserContext).not.toHaveBeenCalled();
    } finally {
      setBrowserLoginInProgress(false);
    }
  });

  it('fails fast optional-token refreshes while a login is in progress', async () => {
    setBrowserLoginInProgress(true);
    try {
      const result = await refreshGraphToken();

      expect(result.ok).toBe(false);
      if (!result.ok) {
        // Optional resources never return AUTH_REQUIRED/AUTH_EXPIRED.
        expect(result.error.code).toBe(ErrorCode.AUTH_INTERACTION_REQUIRED);
      }
      expect(refreshTokensViaHttp).not.toHaveBeenCalled();
    } finally {
      setBrowserLoginInProgress(false);
    }
  });

  it('awaitRefreshQueue also waits for refreshes enqueued while it waits', async () => {
    // Session stays unready so each serialized job reaches the HTTP exchange.
    vi.mocked(extractMessageAuth).mockReturnValue(null);
    vi.mocked(getMessageAuthStatus).mockReturnValue({ hasToken: false });
    let resolveFirst!: (value: Result<void>) => void;
    let resolveSecond!: (value: Result<void>) => void;
    const firstCall = new Promise<Result<void>>(resolve => { resolveFirst = resolve; });
    const secondCall = new Promise<Result<void>>(resolve => { resolveSecond = resolve; });
    vi.mocked(refreshTokensViaHttp)
      .mockImplementationOnce(() => firstCall)
      .mockImplementationOnce(() => secondCall)
      .mockResolvedValue({ ok: true, value: undefined });

    const first = refreshCoreTokensViaHttp();
    let queueSettled = false;
    const queueWait = awaitRefreshQueue().then(() => { queueSettled = true; });
    // Enqueued while the barrier waits: it chains on the same queue tail, so
    // a one-shot snapshot would resolve before this job even starts.
    const second = refreshCoreTokensViaHttp();
    try {
      await new Promise(resolve => setTimeout(resolve, 10));
      expect(queueSettled).toBe(false);

      resolveFirst({ ok: true, value: undefined });
      await new Promise(resolve => setTimeout(resolve, 10));
      // First job settled but the second is still in flight — the barrier
      // must not have resolved yet.
      expect(queueSettled).toBe(false);

      resolveSecond({ ok: true, value: undefined });
      await Promise.allSettled([first, second, queueWait]);
      expect(queueSettled).toBe(true);
    } finally {
      resolveFirst({ ok: true, value: undefined });
      resolveSecond({ ok: true, value: undefined });
      await Promise.allSettled([first, second, queueWait]);
    }
  });

  it('keeps the login guard until the last overlapping login ends', async () => {
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({
      ok: false,
      error: createError(ErrorCode.AUTH_REQUIRED, 'no refresh token'),
    });
    // Two teams_login calls can overlap (handleLogin is not serialized):
    // the first finishing must not reopen the profile under the second.
    setBrowserLoginInProgress(true);
    setBrowserLoginInProgress(true);
    try {
      setBrowserLoginInProgress(false);
      const stillGuarded = await refreshTokensViaBrowser();
      expect(stillGuarded.ok).toBe(false);
      expect(createBrowserContext).not.toHaveBeenCalled();
    } finally {
      setBrowserLoginInProgress(false);
    }

    // Balanced: the guard is released and the browser fallback runs again.
    completeSession();
    vi.mocked(createBrowserContext).mockResolvedValue({ page: {}, context: {} } as never);
    const released = await refreshTokensViaBrowser();
    expect(createBrowserContext).toHaveBeenCalled();
    expect(released.ok).toBe(true);
  });

  it('maps a wedged core HTTP refresh to TIMEOUT at the shared deadline', async () => {
    vi.useFakeTimers();
    // Session is not ready, so the re-check cannot skip the exchange.
    vi.mocked(extractSubstrateToken).mockReturnValue(null);
    vi.mocked(extractMessageAuth).mockReturnValue(null);
    vi.mocked(getMessageAuthStatus).mockReturnValue({ hasToken: false });
    vi.mocked(refreshTokensViaHttp).mockImplementation(() => new Promise(() => {}));

    const pending = refreshCoreTokensViaHttp();
    try {
      await vi.advanceTimersByTimeAsync(CORE_REFRESH_TIMEOUT_MS);

      const result = await pending;
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.error.code).toBe(ErrorCode.TIMEOUT);
        expect(result.error.message).toContain('timed out after 90s');
      }
    } finally {
      // The queue already advanced past the abandoned attempt; settle anyway.
      await Promise.allSettled([pending]);
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
