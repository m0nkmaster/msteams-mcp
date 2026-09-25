import { beforeEach, describe, expect, it, vi } from 'vitest';
import { refreshAssignmentsToken } from './token-refresh.js';
import { refreshTokensViaHttp } from './token-refresh-http.js';
import { extractSubstrateToken, getValidAssignmentsToken } from './token-extractor.js';
import { err, ok } from '../types/result.js';
import { createError, ErrorCode } from '../types/errors.js';

vi.mock('./token-refresh-http.js', () => ({ refreshTokensViaHttp: vi.fn() }));
vi.mock('./token-extractor.js', () => ({ extractSubstrateToken: vi.fn(), getValidAssignmentsToken: vi.fn(), clearTokenCache: vi.fn() }));
vi.mock('../browser/context.js', () => ({ createBrowserContext: vi.fn(async () => ({ page: {}, context: {} })), closeBrowser: vi.fn() }));
vi.mock('../browser/auth.js', () => ({ ensureAuthenticated: vi.fn() }));
const refreshed = ok({ tokensRefreshed: 1, skypeTokenRefreshed: false, refreshTokenRotated: false });

beforeEach(() => {
  vi.resetAllMocks();
  vi.mocked(extractSubstrateToken).mockReturnValue({ token: 'substrate', expiry: new Date(Date.now() + 3600000) });
});
describe('on-demand Assignments refresh', () => {
  it('retries its own exchange after browser recovery instead of treating core success as its token', async () => {
    const expired = err(createError(ErrorCode.AUTH_EXPIRED, 'Expired refresh token'));
    vi.mocked(refreshTokensViaHttp).mockResolvedValueOnce(expired).mockResolvedValueOnce(expired).mockResolvedValueOnce(refreshed);
    vi.mocked(getValidAssignmentsToken).mockReturnValue('assignments-token');
    expect(await refreshAssignmentsToken()).toEqual(ok('assignments-token'));
    expect(vi.mocked(refreshTokensViaHttp).mock.calls).toEqual([['assignments'], [], ['assignments']]);
  });
  it('does not interpret successful exchange without a valid audience as feature unavailability', async () => {
    vi.mocked(refreshTokensViaHttp).mockResolvedValue(refreshed);
    vi.mocked(getValidAssignmentsToken).mockReturnValue(null);
    expect(await refreshAssignmentsToken()).toMatchObject({ ok: false, error: { code: ErrorCode.API_ERROR } });
  });
  it('does not run core refresh on an access refusal', async () => {
    const denied = err(createError(ErrorCode.ACCESS_DENIED, 'Consent required'));
    vi.mocked(refreshTokensViaHttp).mockResolvedValue(denied);
    expect(await refreshAssignmentsToken()).toEqual(denied);
    expect(refreshTokensViaHttp).toHaveBeenCalledTimes(1);
  });
  it('does not trigger another generic login when only Assignments still requires authorization', async () => {
    const challenge = err(createError(ErrorCode.AUTH_EXPIRED, 'MFA required for Assignments'));
    vi.mocked(refreshTokensViaHttp).mockResolvedValueOnce(challenge).mockResolvedValueOnce(refreshed).mockResolvedValueOnce(challenge);
    expect(await refreshAssignmentsToken()).toMatchObject({ ok: false, error: { code: ErrorCode.AUTH_INTERACTION_REQUIRED, retryable: false } });
  });

});
