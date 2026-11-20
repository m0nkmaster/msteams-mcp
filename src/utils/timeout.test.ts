import { describe, expect, it, vi, afterEach } from 'vitest';
import { TimeoutError, withTimeout } from './timeout.js';

afterEach(() => {
  vi.useRealTimers();
});

describe('withTimeout', () => {
  it('resolves with the promise value when it settles in time', async () => {
    await expect(withTimeout(Promise.resolve('done'), 1000, 'op')).resolves.toBe('done');
  });

  it('rejects with the promise error when it fails in time', async () => {
    await expect(withTimeout(Promise.reject(new Error('boom')), 1000, 'op')).rejects.toThrow('boom');
  });

  it('rejects with TimeoutError when the promise never settles', async () => {
    vi.useFakeTimers();
    const pending = withTimeout(new Promise(() => {}), 5000, 'stuck op');
    const assertion = expect(pending).rejects.toBeInstanceOf(TimeoutError);
    await vi.advanceTimersByTimeAsync(5000);
    await assertion;
  });

  it('clears its timer after the promise settles', async () => {
    vi.useFakeTimers();
    await expect(withTimeout(Promise.resolve(1), 5000, 'op')).resolves.toBe(1);
    // Advancing past the deadline must not surface a late rejection.
    await vi.advanceTimersByTimeAsync(10000);
  });
});
