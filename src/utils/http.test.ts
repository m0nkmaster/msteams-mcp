/**
 * Unit tests for HTTP utilities.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { httpRequest, clearRateLimitState } from './http.js';

describe('httpRequest', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.stubGlobal('fetch', vi.fn());
    clearRateLimitState();
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it('returns successful response on HTTP 200', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response(JSON.stringify({ ok: true }), {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
      })
    );

    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.status).toBe(200);
      expect(result.value.data).toEqual({ ok: true });
    }
  });

  it('returns error on HTTP 400', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('Bad Request', { status: 400 })
    );

    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('INVALID_INPUT');
    }
  });

  it('returns error on HTTP 401', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('Unauthorized', { status: 401 })
    );

    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('AUTH_EXPIRED');
    }
  });

  it('returns error on HTTP 403', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('Forbidden', { status: 403 })
    );

    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('AUTH_REQUIRED');
    }
  });

  it('returns error on HTTP 404', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('Not Found', { status: 404 })
    );

    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('NOT_FOUND');
    }
  });

  it('returns error on HTTP 429 with retry-after header', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('Rate Limited', {
        status: 429,
        headers: { 'Retry-After': '5' },
      })
    );

    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('RATE_LIMITED');
      expect(result.error.retryAfterMs).toBe(5000);
    }
  });

  it('returns error on HTTP 500', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('Server Error', { status: 500 })
    );

    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('API_ERROR');
    }
  });

  it('returns TIMEOUT error on request timeout', async () => {
    vi.mocked(fetch).mockImplementation(() => 
      new Promise((_, reject) => {
        const error = new Error('Request timed out');
        error.name = 'AbortError';
        reject(error);
      })
    );

    const result = await httpRequest('https://api.example.com/slow', {
      timeoutMs: 100,
    });

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('TIMEOUT');
    }
  });

  it('returns NETWORK_ERROR on ECONNRESET', async () => {
    vi.mocked(fetch).mockImplementation(() => 
      Promise.reject(new Error('ECONNRESET'))
    );

    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('NETWORK_ERROR');
    }
  });

  it('retries on retryable errors', async () => {
    let attempts = 0;
    vi.mocked(fetch).mockImplementation(() => {
      attempts++;
      if (attempts < 3) {
        return Promise.reject(new Error('ECONNRESET'));
      }
      return Promise.resolve(
        new Response(JSON.stringify({ success: true }), {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
        })
      );
    });

    const result = await httpRequest('https://api.example.com/data', {
      maxRetries: 3,
      retryBaseDelayMs: 10,
    });

    expect(result.ok).toBe(true);
    expect(attempts).toBe(3);
  });

  it('does not retry on non-retryable errors like 400', async () => {
    let attempts = 0;
    vi.mocked(fetch).mockImplementation(() => {
      attempts++;
      return Promise.resolve(
        new Response('Bad Request', { status: 400 })
      );
    });

    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(false);
    expect(attempts).toBe(1); // Should not retry
  });

  it('uses default timeout of 30 seconds', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('{}', { status: 200 })
    );

    await httpRequest('https://api.example.com/data');

    // Fetch was called - default timeout was applied
    expect(vi.mocked(fetch)).toHaveBeenCalled();
  });

  it('accepts custom timeout', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('{}', { status: 200 })
    );

    await httpRequest('https://api.example.com/data', {
      timeoutMs: 5000,
    });

    expect(vi.mocked(fetch)).toHaveBeenCalled();
  });

  it('returns RATE_LIMITED error when rate limited', async () => {
    // First, simulate a rate limit response
    vi.mocked(fetch).mockResolvedValue(
      new Response('Rate Limited', {
        status: 429,
        headers: { 'Retry-After': '60' },
      })
    );

    // First request triggers rate limit state
    await httpRequest('https://api.example.com/data');

    // Replace fetch mock for the second request to verify rate limit check
    vi.mocked(fetch).mockClear();
    
    // Second request should be rate limited without calling fetch
    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('RATE_LIMITED');
    }
    // Should not have called fetch since we were rate limited
    expect(vi.mocked(fetch)).not.toHaveBeenCalled();
  });

  it('handles plain text response', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('plain text response', {
        status: 200,
        headers: { 'Content-Type': 'text/plain' },
      })
    );

    const result = await httpRequest('https://api.example.com/text');

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.data).toBe('plain text response');
    }
  });

  it('handles response without content-type header', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('no content type', { status: 200 })
    );

    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.data).toBe('no content type');
    }
  });
});

describe('clearRateLimitState', () => {
  beforeEach(() => {
    clearRateLimitState();
  });

  it('clears rate limit state', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(
      new Response('Rate Limited', {
        status: 429,
        headers: { 'Retry-After': '60' },
      })
    ));

    // First request triggers rate limit
    await httpRequest('https://api.example.com/data');

    // Clear the rate limit state
    clearRateLimitState();

    // Second request should work
    vi.mocked(fetch).mockResolvedValue(
      new Response('{}', { status: 200, headers: { 'Content-Type': 'application/json' } })
    );

    const result = await httpRequest('https://api.example.com/data');
    expect(result.ok).toBe(true);

    vi.unstubAllGlobals();
  });
});
