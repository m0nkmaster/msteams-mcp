import { beforeEach, describe, expect, it, vi } from 'vitest';
import { extractAssignmentsToken, extractGraphToken } from '../auth/token-extractor.js';
import { refreshAssignmentsToken, refreshGraphToken } from '../auth/token-refresh.js';
import { requireAssignmentsTokenAsync, requireGraphTokenAsync, resetAssignmentsAvailability } from './auth-guards.js';
import { ErrorCode, createError } from '../types/errors.js';
import { err, ok } from '../types/result.js';

vi.mock('../auth/token-extractor.js', () => ({ extractAssignmentsToken: vi.fn(), extractGraphToken: vi.fn() }));
vi.mock('../auth/token-refresh.js', () => ({ refreshAssignmentsToken: vi.fn(), refreshGraphToken: vi.fn() }));

beforeEach(() => {
  vi.resetAllMocks();
  resetAssignmentsAvailability();
  vi.mocked(extractAssignmentsToken).mockReturnValue(null);
  vi.mocked(extractGraphToken).mockReturnValue(null);
});

describe('Assignments auth', () => {
  it('remembers definitive access refusals and resets them on login', async () => {
    const denied = err(createError(ErrorCode.ACCESS_DENIED, 'Consent required'));
    vi.mocked(refreshAssignmentsToken).mockResolvedValue(denied);
    expect(await requireAssignmentsTokenAsync()).toEqual(denied);
    expect(await requireAssignmentsTokenAsync()).toEqual(denied);
    expect(refreshAssignmentsToken).toHaveBeenCalledTimes(1);
    resetAssignmentsAvailability();
    await requireAssignmentsTokenAsync();
    expect(refreshAssignmentsToken).toHaveBeenCalledTimes(2);
  });

  it.each([ErrorCode.NETWORK_ERROR, ErrorCode.AUTH_EXPIRED, ErrorCode.API_ERROR, ErrorCode.AUTH_INTERACTION_REQUIRED])('does not cache %s as unavailable', async code => {
    const failure = err(createError(code, 'Temporary failure'));
    vi.mocked(refreshAssignmentsToken).mockResolvedValue(failure);
    expect(await requireAssignmentsTokenAsync()).toEqual(failure);
    await requireAssignmentsTokenAsync();
    expect(refreshAssignmentsToken).toHaveBeenCalledTimes(2);
  });

  it('keeps Graph availability independent of Assignments', async () => {
    vi.mocked(refreshAssignmentsToken).mockResolvedValue(err(createError(ErrorCode.ACCESS_DENIED, 'No Assignments')));
    vi.mocked(refreshGraphToken).mockResolvedValue(ok('graph-token'));
    await requireAssignmentsTokenAsync();
    expect(await requireGraphTokenAsync()).toEqual(ok('graph-token'));
    vi.mocked(refreshGraphToken).mockResolvedValue(err(createError(ErrorCode.ACCESS_DENIED, 'No Graph')));
    await requireGraphTokenAsync();
    await requireGraphTokenAsync();
    expect(refreshGraphToken).toHaveBeenCalledTimes(2);
    resetAssignmentsAvailability();
    await requireGraphTokenAsync();
    expect(refreshGraphToken).toHaveBeenCalledTimes(3);
  });

  it('coalesces simultaneous requests for a missing token', async () => {
    vi.mocked(refreshAssignmentsToken).mockResolvedValue(ok('new-token'));
    expect(await Promise.all([requireAssignmentsTokenAsync(), requireAssignmentsTokenAsync()]))
      .toEqual([ok('new-token'), ok('new-token')]);
    expect(refreshAssignmentsToken).toHaveBeenCalledTimes(1);
  });

  it('does not refresh a healthy token', async () => {
    vi.mocked(extractAssignmentsToken).mockReturnValue({ token: 'valid', expiry: new Date(Date.now() + 3600000) });
    expect(await requireAssignmentsTokenAsync()).toEqual(ok('valid'));
    expect(refreshAssignmentsToken).not.toHaveBeenCalled();
  });
});
