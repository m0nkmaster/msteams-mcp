import { describe, it, expect, vi, beforeEach } from 'vitest';
import { httpRequest } from '../utils/http.js';
import { requireGraphMailTokenAsync } from '../utils/auth-guards.js';
import { invalidateAccessToken } from '../auth/token-extractor.js';
import { ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import { getEmail } from './graph-mail-api.js';
import { graphEmailMessage } from '../__fixtures__/api-responses.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({ requireGraphMailTokenAsync: vi.fn() }));
vi.mock('../auth/token-extractor.js', () => ({ invalidateAccessToken: vi.fn() }));

const REST_ID = 'AAMkAGE1OWFlZjc0.RestId-Q3Budget';
const EWS_ID = 'AAMkEws+Format/Id=';

const mockHttp = vi.mocked(httpRequest);

function message(overrides: Record<string, unknown> = {}) {
  return ok({ status: 200, headers: new Headers(), data: { ...graphEmailMessage, ...overrides } });
}

function translated(targetId: string | undefined) {
  return ok({ status: 200, headers: new Headers(), data: { value: [{ sourceId: EWS_ID, targetId }] } });
}

beforeEach(() => {
  vi.clearAllMocks();
  vi.mocked(requireGraphMailTokenAsync).mockResolvedValue(ok('graph-token'));
});

describe('getEmail', () => {
  it('fetches a message by id and parses it', async () => {
    mockHttp.mockResolvedValue(message());

    const result = await getEmail(REST_ID);

    expect(result).toMatchObject({ ok: true, value: { id: REST_ID, subject: 'Q3 Budget Review', bodyFormat: 'text' } });
    const [url, opts] = mockHttp.mock.calls[0];
    const parsed = new URL(url as string);
    expect(parsed.origin).toBe('https://graph.microsoft.com');
    expect(parsed.pathname).toBe(`/v1.0/me/messages/${encodeURIComponent(REST_ID)}`);
    expect(parsed.searchParams.get('$select')).toContain('internetMessageId');
    expect(parsed.searchParams.get('$expand')).toContain('attachments');
    expect((opts as { headers: Record<string, string> }).headers.Authorization).toBe('Bearer graph-token');
  });

  it('translates an EWS-format id and retries when the direct get fails', async () => {
    mockHttp
      .mockResolvedValueOnce(err(createError(ErrorCode.INVALID_INPUT, 'HTTP 400')))
      .mockResolvedValueOnce(translated(REST_ID))
      .mockResolvedValueOnce(message());

    const result = await getEmail(EWS_ID);

    expect(result).toMatchObject({ ok: true, value: { subject: 'Q3 Budget Review' } });
    const translateCall = mockHttp.mock.calls[1];
    expect(translateCall[0]).toBe('https://graph.microsoft.com/v1.0/me/translateExchangeIds');
    expect((translateCall[1] as { body: string }).body).toContain('"sourceIdType":"ewsId"');
  });

  it('reports not found when translation returns the same id', async () => {
    mockHttp
      .mockResolvedValueOnce(err(createError(ErrorCode.NOT_FOUND, 'HTTP 404')))
      .mockResolvedValueOnce(translated(EWS_ID));

    const result = await getEmail(EWS_ID);

    expect(result).toMatchObject({ ok: false, error: { code: ErrorCode.NOT_FOUND, retryable: false } });
    expect(mockHttp).toHaveBeenCalledTimes(2);
  });

  it('passes through a translation failure with its own code and retryability', async () => {
    mockHttp
      .mockResolvedValueOnce(err(createError(ErrorCode.NOT_FOUND, 'HTTP 404')))
      .mockResolvedValueOnce(err(createError(ErrorCode.RATE_LIMITED, 'HTTP 429', { retryable: true })));

    expect(await getEmail(EWS_ID)).toMatchObject({ ok: false, error: { code: ErrorCode.RATE_LIMITED, retryable: true } });
  });

  it('discards the token when translation is rejected', async () => {
    mockHttp
      .mockResolvedValueOnce(err(createError(ErrorCode.NOT_FOUND, 'HTTP 404')))
      .mockResolvedValueOnce(err(createError(ErrorCode.AUTH_EXPIRED, 'HTTP 401')));

    expect(await getEmail(EWS_ID)).toMatchObject({ ok: false, error: { code: ErrorCode.API_ERROR, retryable: true } });
    expect(invalidateAccessToken).toHaveBeenCalledWith('graph-token');
  });

  it('does not translate a REST id that is not found', async () => {
    mockHttp.mockResolvedValueOnce(err(createError(ErrorCode.NOT_FOUND, 'HTTP 404: ErrorItemNotFound')));

    const result = await getEmail(REST_ID);

    expect(result).toMatchObject({ ok: false, error: { code: ErrorCode.NOT_FOUND } });
    expect(result.ok || result.error.message).toContain('ErrorItemNotFound');
    expect(mockHttp).toHaveBeenCalledTimes(1);
  });

  it("keeps Graph's message on a 400 instead of reporting not found", async () => {
    mockHttp.mockResolvedValueOnce(err(createError(ErrorCode.INVALID_INPUT, 'HTTP 400: Invalid $expand')));

    expect(await getEmail(REST_ID)).toMatchObject({ ok: false, error: { code: ErrorCode.INVALID_INPUT, message: 'HTTP 400: Invalid $expand' } });
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
    vi.mocked(requireGraphMailTokenAsync).mockResolvedValue(unavailable);

    expect(await getEmail(REST_ID)).toEqual(unavailable);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it.each(['', 'has space', 'has\ttab'])('rejects an invalid id before using the token: %s', async id => {
    expect(await getEmail(id)).toMatchObject({ ok: false, error: { code: ErrorCode.INVALID_INPUT } });
    expect(requireGraphMailTokenAsync).not.toHaveBeenCalled();
  });

  it('passes the requested format to the parser', async () => {
    mockHttp.mockResolvedValue(message());

    const result = await getEmail(REST_ID, { format: 'html' });

    expect(result).toMatchObject({ ok: true, value: { bodyFormat: 'html' } });
  });
});
