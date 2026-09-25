/**
 * Tests for requireAssignmentsTokenAsync: Assignments is optional, so accounts
 * without it must get a clear non-auth error and no repeated refreshes.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ErrorCode, createError } from '../types/errors.js';
import { ok, err } from '../types/result.js';

vi.mock('../auth/token-extractor.js', () => ({
  getValidSubstrateToken: vi.fn(),
  getValidAssignmentsToken: vi.fn(),
  extractAssignmentsToken: vi.fn(),
  extractMessageAuth: vi.fn(),
  extractCsaToken: vi.fn(),
  extractSubstrateToken: vi.fn(),
  extractSkypeSpacesToken: vi.fn(),
  extractRegionConfig: vi.fn(),
  getUserProfile: vi.fn(),
  clearTokenCache: vi.fn(),
}));
vi.mock('../auth/token-refresh.js', () => ({ refreshAssignmentsToken: vi.fn() }));

import { getValidAssignmentsToken, extractAssignmentsToken } from '../auth/token-extractor.js';
import { refreshAssignmentsToken } from '../auth/token-refresh.js';
import { requireAssignmentsTokenAsync, resetAssignmentsAvailability } from './auth-guards.js';

const unavailable = () => err(createError(ErrorCode.ACCESS_DENIED, 'Assignments not available: consent required', { retryable: false }));

describe('requireAssignmentsTokenAsync', () => {
  beforeEach(() => {
    vi.resetAllMocks();
    resetAssignmentsAvailability();
    vi.mocked(extractAssignmentsToken).mockReturnValue(null);
    vi.mocked(getValidAssignmentsToken).mockReturnValue(null);
  });

  it('returns a non-retryable, non-auth error when the account has no Assignments', async () => {
    vi.mocked(refreshAssignmentsToken).mockResolvedValue(unavailable());

    const result = await requireAssignmentsTokenAsync();

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe(ErrorCode.ACCESS_DENIED);
      expect(result.error.retryable).toBe(false);
      expect(result.error.message).toContain('not available');
    }
  });

  it('does not refresh again once Assignments is known to be unavailable', async () => {
    vi.mocked(refreshAssignmentsToken).mockResolvedValue(unavailable());

    await requireAssignmentsTokenAsync();
    await requireAssignmentsTokenAsync();
    await requireAssignmentsTokenAsync();

    expect(refreshAssignmentsToken).toHaveBeenCalledTimes(1);
  });

  it('does not remember unavailability unless Azure AD definitively refused the scope', async () => {
    vi.mocked(refreshAssignmentsToken).mockResolvedValue(err(createError(ErrorCode.API_ERROR, 'No valid token returned', { retryable: false })));

    const first = await requireAssignmentsTokenAsync();
    await requireAssignmentsTokenAsync();

    expect(first.ok).toBe(false);
    if (!first.ok) expect(first.error.code).toBe(ErrorCode.API_ERROR);
    expect(refreshAssignmentsToken).toHaveBeenCalledTimes(2);
  });

  it('returns the token minted by the refresh', async () => {
    vi.mocked(refreshAssignmentsToken).mockResolvedValue(ok('assignments-jwt'));
    expect(await requireAssignmentsTokenAsync()).toEqual(ok('assignments-jwt'));
  });

  it('skips refresh when a valid token has plenty of time left', async () => {
    vi.mocked(extractAssignmentsToken).mockReturnValue({
      token: 'assignments-jwt',
      expiry: new Date(Date.now() + 60 * 60 * 1000),
    });
    vi.mocked(getValidAssignmentsToken).mockReturnValue('assignments-jwt');

    const result = await requireAssignmentsTokenAsync();

    expect(result).toEqual({ ok: true, value: 'assignments-jwt' });
    expect(refreshAssignmentsToken).not.toHaveBeenCalled();
  });

  it('keeps AUTH_EXPIRED for genuine refresh failures', async () => {
    vi.mocked(refreshAssignmentsToken).mockResolvedValue(
      err(createError(ErrorCode.AUTH_EXPIRED, 'expired'))
    );

    const result = await requireAssignmentsTokenAsync();

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe(ErrorCode.AUTH_EXPIRED);
  });

  it('surfaces transient refresh failures without telling the user to log in', async () => {
    vi.mocked(refreshAssignmentsToken).mockResolvedValue(
      err(createError(ErrorCode.UNKNOWN, 'Token refresh already in progress.', { retryable: true }))
    );

    const result = await requireAssignmentsTokenAsync();

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe(ErrorCode.UNKNOWN);
      expect(result.error.retryable).toBe(true);
    }
  });
});
