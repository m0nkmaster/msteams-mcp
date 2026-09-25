import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { mkdtemp, readFile, rm, writeFile, stat, open } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { downloadFile } from './download-api.js';
import { getSharePointToken } from '../auth/token-refresh-http.js';
import { clearRateLimitState } from '../utils/http.js';
import { ok } from '../types/result.js';
import { createHash } from 'node:crypto';
import { createReadStream } from 'node:fs';

vi.mock('../auth/token-refresh-http.js', () => ({ getSharePointToken: vi.fn() }));
const url = 'https://example-my.sharepoint.com/personal/test/Documents/file.md';
let directory: string;
beforeEach(async () => {
  directory = await mkdtemp(join(tmpdir(), 'teams-download-'));
  vi.mocked(getSharePointToken).mockReset().mockResolvedValue(ok('test-token'));
  clearRateLimitState();
});
afterEach(async () => {
  vi.useRealTimers();
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
  await rm(directory, { recursive: true, force: true });
});

describe('downloadFile', () => {
  it('preserves binary bytes, even with a JSON content type', async () => {
    const bytes = Buffer.from([0, 255, 254, 128, 10]);
    const fetchMock = vi.fn().mockResolvedValue(new Response(bytes, { headers: { 'content-type': 'application/json' } }));
    vi.stubGlobal('fetch', fetchMock);
    const output = join(directory, 'file');
    const result = await downloadFile(url, output);
    expect(result.ok).toBe(true);
    expect(await readFile(output)).toEqual(bytes);
    expect(fetchMock).toHaveBeenCalledWith(expect.stringContaining('/_api/web/GetFileByServerRelativePath('), expect.objectContaining({
      redirect: 'error', headers: { Authorization: 'Bearer test-token' },
    }));
  });
  it('does not overwrite an existing file', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response('new')));
    const output = join(directory, 'file');
    await writeFile(output, 'original');
    expect((await downloadFile(url, output)).ok).toBe(false);
    expect(await readFile(output, 'utf8')).toBe('original');
  });
  it('rejects unsupported URLs before accessing credentials', async () => {
    expect((await downloadFile('https://example.org/file.md', join(directory, 'file'))).ok).toBe(false);
    expect(getSharePointToken).not.toHaveBeenCalled();
  });
  it('rejects relative output paths', async () => {
    expect((await downloadFile(url, 'file.md')).ok).toBe(false);
    expect(getSharePointToken).not.toHaveBeenCalled();
  });
  it('refreshes the host token once after a 401', async () => {
    const fetchMock = vi.fn().mockResolvedValueOnce(new Response('expired', { status: 401 }))
      .mockResolvedValueOnce(new Response('# markdown'));
    vi.stubGlobal('fetch', fetchMock);
    expect((await downloadFile(url, join(directory, 'file'))).ok).toBe(true);
    expect(getSharePointToken).toHaveBeenLastCalledWith('https://example-my.sharepoint.com', true);
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });
  it('does not create a file when the server denies access', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response('forbidden', { status: 403 })));
    const output = join(directory, 'file');
    expect((await downloadFile(url, output)).ok).toBe(false);
    await expect(readFile(output)).rejects.toThrow();
  });
  it('streams 64 MiB to disk before the response finishes and hashes all chunks', async () => {
    const output = join(directory, 'large-file');
    const chunk = Buffer.alloc(1024 * 1024, 0xa5);
    const expected = createHash('sha256');
    let sent = 0;
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response(new ReadableStream({
      async pull(controller) {
        if (sent === 4) {
          // The consumer must write before requesting the rest, not buffer to EOF.
          expect((await stat(output)).size).toBeGreaterThanOrEqual(2 * chunk.length);
        }
        if (sent === 64) {
          controller.close();
        } else {
          expected.update(chunk);
          controller.enqueue(chunk);
          sent++;
        }
      },
    }))));
    const result = await downloadFile(url, output);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.size).toBe(64 * chunk.length);
    expect((await stat(output)).size).toBe(result.value.size);
    expect(result.value.sha256).toBe(expected.digest('hex'));
    const actual = createHash('sha256');
    for await (const bytes of createReadStream(output)) actual.update(bytes);
    expect(result.value.sha256).toBe(actual.digest('hex'));
  });

  it('removes a partially written file after a network interruption without replaying the request', async () => {
    let sent = 0;
    const fetchMock = vi.fn().mockImplementation(() => new Response(new ReadableStream({
      pull(controller) {
        if (sent++ < 3) controller.enqueue(Buffer.from('partial data'));
        else controller.error(new Error('ECONNRESET'));
      },
    })));
    vi.stubGlobal('fetch', fetchMock);
    const output = join(directory, 'file');
    const result = await downloadFile(url, output);
    expect(result.ok).toBe(false);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    await expect(stat(output)).rejects.toThrow();
  });

  it('cancels the response and removes the partial file if writing fails', async () => {
    const output = join(directory, 'file');
    // Node's FileHandle prototype is shared by handles opened by the downloader.
    const probe = await open(join(directory, 'probe'), 'wx');
    const prototype = Object.getPrototypeOf(probe);
    await probe.close();
    vi.spyOn(prototype, 'writeFile').mockRejectedValueOnce(new Error('ENOSPC'));
    const cancel = vi.fn();
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response(new ReadableStream({
      pull(controller) { controller.enqueue(Buffer.from('data')); },
      cancel,
    }))));
    const result = await downloadFile(url, output);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.message).toContain('ENOSPC');
    expect(cancel).toHaveBeenCalled();
    await expect(stat(output)).rejects.toThrow();
  });

  it('removes the partial file when a response stalls', async () => {
    let started!: () => void;
    const ready = new Promise<void>(resolve => { started = resolve; });
    vi.stubGlobal('fetch', vi.fn().mockImplementation((_url, options) => {
      vi.useFakeTimers();
      const response = new Response(new ReadableStream({
        start(controller) {
          controller.enqueue(Buffer.from('partial'));
          options.signal.addEventListener('abort', () => controller.error(options.signal.reason), { once: true });
        },
      }));
      started();
      return response;
    }));
    const output = join(directory, 'stalled');
    const pending = downloadFile(url, output);
    await ready;
    // Let the first chunk finish writing before advancing the inactivity timer.
    await vi.waitFor(async () => expect((await stat(output)).size).toBe(7));
    await vi.advanceTimersByTimeAsync(30001);
    const result = await pending;
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe('TIMEOUT');
    await expect(stat(output)).rejects.toThrow();
    expect(vi.getTimerCount()).toBe(0);
  });

  it('saves an empty file with the empty-content hash', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response(null)));
    const output = join(directory, 'empty');
    const result = await downloadFile(url, output);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.size).toBe(0);
      expect(result.value.sha256).toBe(createHash('sha256').digest('hex'));
    }
    expect((await stat(output)).size).toBe(0);
  });
});
