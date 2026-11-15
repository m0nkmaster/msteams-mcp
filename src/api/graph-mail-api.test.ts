import { describe, it, expect, vi, beforeEach } from 'vitest';
import { httpRequest } from '../utils/http.js';
import { requireGraphTokenAsync } from '../utils/auth-guards.js';
import { invalidateAccessToken } from '../auth/token-extractor.js';
import { ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import { getEmail } from './graph-mail-api.js';
import { graphEmailMessage } from '../__fixtures__/api-responses.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({ requireGraphTokenAsync: vi.fn() }));
vi.mock('../auth/token-extractor.js', () => ({ invalidateAccessToken: vi.fn() }));

const REST_ID = 'AAMkAGE1OWFlZjc0.RestId-Q3Budget';
const mockHttp = vi.mocked(httpRequest);

beforeEach(() => {
  vi.clearAllMocks();
  vi.mocked(requireGraphTokenAsync).mockResolvedValue(ok('graph-token'));
});

describe('getEmail', () => {
  it('fetches a message by id and parses it', async () => {
    mockHttp.mockResolvedValue(ok({ status: 200, headers: new Headers(), data: graphEmailMessage }));

    const result = await getEmail(REST_ID);

    expect(result).toMatchObject({ ok: true, value: { id: REST_ID, subject: 'Q3 Budget Review' } });
    const [url, opts] = mockHttp.mock.calls[0];
    const parsed = new URL(url as string);
    expect(parsed.origin).toBe('https://graph.microsoft.com');
    expect(parsed.pathname).toBe(`/v1.0/me/messages/${encodeURIComponent(REST_ID)}`);
    expect(parsed.searchParams.get('$select')).toContain('internetMessageId');
    expect(parsed.searchParams.get('$expand')).toContain('attachments');
    expect((opts as { headers: Record<string, string> }).headers.Authorization).toBe('Bearer graph-token');
  });

  it("reports not found with Graph's message", async () => {
    mockHttp.mockResolvedValue(err(createError(ErrorCode.NOT_FOUND, 'HTTP 404: ErrorItemNotFound')));

    const result = await getEmail(REST_ID);

    expect(result).toMatchObject({ ok: false, error: { code: ErrorCode.NOT_FOUND, retryable: false } });
    expect(!result.ok && result.error.message).toContain('ErrorItemNotFound');
  });

  it('passes other errors through unchanged', async () => {
    const throttled = err(createError(ErrorCode.RATE_LIMITED, 'HTTP 429', { retryable: true }));
    mockHttp.mockResolvedValue(throttled);

    expect(await getEmail(REST_ID)).toEqual(throttled);
  });

  it('discards a rejected Graph token without triggering Teams re-login', async () => {
    mockHttp.mockResolvedValue(err(createError(ErrorCode.AUTH_EXPIRED, 'HTTP 401')));

    const result = await getEmail(REST_ID);

    expect(result).toMatchObject({ ok: false, error: { code: ErrorCode.API_ERROR, retryable: true } });
    expect(invalidateAccessToken).toHaveBeenCalledWith('graph-token');
  });

  it('reports a forbidden token as access denied, not an auth error', async () => {
    mockHttp.mockResolvedValue(err(createError(ErrorCode.AUTH_REQUIRED, 'HTTP 403')));

    expect(await getEmail(REST_ID)).toMatchObject({ ok: false, error: { code: ErrorCode.ACCESS_DENIED } });
  });

  it('passes through an unavailable Graph token unchanged', async () => {
    const unavailable = err(createError(ErrorCode.AUTH_INTERACTION_REQUIRED, 'Microsoft Graph could not be authorized'));
    vi.mocked(requireGraphTokenAsync).mockResolvedValue(unavailable);

    expect(await getEmail(REST_ID)).toEqual(unavailable);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it.each(['', 'has space', 'has\ttab'])('rejects an invalid id before using the token: %s', async id => {
    expect(await getEmail(id)).toMatchObject({ ok: false, error: { code: ErrorCode.INVALID_INPUT } });
    expect(requireGraphTokenAsync).not.toHaveBeenCalled();
  });
});
