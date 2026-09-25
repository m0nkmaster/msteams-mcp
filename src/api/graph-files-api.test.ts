import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createHash } from 'node:crypto';
import { httpRequest } from '../utils/http.js';
import { requireGraphTokenAsync } from '../utils/auth-guards.js';
import { invalidateAccessToken } from '../auth/token-extractor.js';
import { ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import { downloadDriveItem } from './graph-files-api.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({ requireGraphTokenAsync: vi.fn() }));
vi.mock('../auth/token-extractor.js', () => ({ invalidateAccessToken: vi.fn() }));

const FILE_URL = 'https://graph.microsoft.com/v1.0/drives/b!abc_DEF-123/items/01VO3D63DYJHLB2KJCVFD3YH77DMC5RIEJ';
const DOWNLOAD_URL = 'https://school.sharepoint.com/sites/x/_layouts/15/download.aspx?UniqueId=1&tempauth=secret';
const BYTES = Buffer.from('PK\u0003\u0004 pretend pptx bytes');

const mockHttp = vi.mocked(httpRequest);
const fetchMock = vi.fn();
let dir: string;

function item(overrides: Record<string, unknown> = {}) {
  return ok({ status: 200, headers: new Headers(), data: {
    name: 'Lesson.pptx', size: BYTES.length, file: { mimeType: 'application/vnd.ms-powerpoint' },
    currentUserRole: { blocksDownload: false }, '@microsoft.graph.downloadUrl': DOWNLOAD_URL, ...overrides,
  } });
}

beforeEach(async () => {
  vi.clearAllMocks();
  vi.mocked(requireGraphTokenAsync).mockResolvedValue(ok('graph-token'));
  fetchMock.mockImplementation(async () => new Response(BYTES, { status: 200, headers: { 'content-type': 'application/octet-stream' } }));
  vi.stubGlobal('fetch', fetchMock);
  dir = await mkdtemp(join(tmpdir(), 'graph-files-'));
});

afterEach(async () => {
  vi.unstubAllGlobals();
  await rm(dir, { recursive: true, force: true });
});

describe('downloadDriveItem', () => {
  it('streams the file to disk with owner-only permissions and reports its hash', async () => {
    mockHttp.mockResolvedValue(item());
    const out = join(dir, 'lesson.pptx');

    const result = await downloadDriveItem(FILE_URL, out);

    expect(result).toEqual(ok({
      name: 'Lesson.pptx', outputPath: out, size: BYTES.length, contentType: 'application/octet-stream',
      sha256: createHash('sha256').update(BYTES).digest('hex'),
    }));
    expect(await readFile(out)).toEqual(BYTES);
    expect((await stat(out)).mode & 0o777).toBe(0o600);
    const [metaUrl, metaOpts] = mockHttp.mock.calls[0];
    expect(new URL(metaUrl as string).searchParams.get('$select')).toContain('content.downloadUrl');
    expect((metaOpts as { headers: Record<string, string> }).headers.Authorization).toBe('Bearer graph-token');
    // The pre-authenticated link gets no bearer token and no redirects.
    const [dlUrl, dlOpts] = fetchMock.mock.calls[0];
    expect(dlUrl).toBe(DOWNLOAD_URL);
    expect(dlOpts).toMatchObject({ redirect: 'error' });
    expect((dlOpts as RequestInit).headers).toBeUndefined();
  });

  it.each([
    'https://evil.example/v1.0/drives/b!abc/items/01X',
    'https://graph.microsoft.com/v1.0/me/drive/items/01X',
    `${FILE_URL}?$select=x`,
    `${FILE_URL}/../../me`,
  ])('rejects anything but a Graph drive-item URL before using the token: %s', async url => {
    expect(await downloadDriveItem(url, join(dir, 'x'))).toMatchObject({ ok: false, error: { code: ErrorCode.INVALID_INPUT } });
    expect(requireGraphTokenAsync).not.toHaveBeenCalled();
  });

  it('requires an absolute output path', async () => {
    expect(await downloadDriveItem(FILE_URL, 'relative.pptx')).toMatchObject({ ok: false, error: { code: ErrorCode.INVALID_INPUT } });
    expect(requireGraphTokenAsync).not.toHaveBeenCalled();
  });

  it('never overwrites an existing file', async () => {
    mockHttp.mockResolvedValue(item());
    const out = join(dir, 'existing.pptx');
    await writeFile(out, 'keep me');
    expect(await downloadDriveItem(FILE_URL, out)).toMatchObject({ ok: false, error: { code: ErrorCode.INVALID_INPUT } });
    expect(await readFile(out, 'utf8')).toBe('keep me');
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('respects an owner download block without creating a file', async () => {
    mockHttp.mockResolvedValue(item({ currentUserRole: { blocksDownload: true } }));
    const out = join(dir, 'blocked.pptx');
    expect(await downloadDriveItem(FILE_URL, out)).toMatchObject({ ok: false, error: { code: ErrorCode.ACCESS_DENIED } });
    expect(existsSync(out)).toBe(false);
  });

  it.each(['https://evil.example/file', 'http://school.sharepoint.com/file', undefined])('refuses a non-SharePoint download link: %s', async link => {
    mockHttp.mockResolvedValue(item({ '@microsoft.graph.downloadUrl': link }));
    expect(await downloadDriveItem(FILE_URL, join(dir, 'x'))).toMatchObject({ ok: false, error: { code: ErrorCode.API_ERROR } });
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('rejects folders', async () => {
    mockHttp.mockResolvedValue(item({ file: undefined }));
    expect(await downloadDriveItem(FILE_URL, join(dir, 'x'))).toMatchObject({ ok: false, error: { code: ErrorCode.INVALID_INPUT } });
  });

  it('removes a truncated download', async () => {
    mockHttp.mockResolvedValue(item({ size: BYTES.length + 10 }));
    const out = join(dir, 'short.pptx');
    expect(await downloadDriveItem(FILE_URL, out)).toMatchObject({ ok: false, error: { code: ErrorCode.API_ERROR } });
    expect(existsSync(out)).toBe(false);
  });

  it('removes the partial file when the stream fails', async () => {
    mockHttp.mockResolvedValue(item());
    fetchMock.mockImplementation(async () => new Response(new ReadableStream({
      start(controller) { controller.enqueue(new Uint8Array([1, 2, 3])); controller.error(new Error('connection reset')); },
    })));
    const out = join(dir, 'broken.pptx');
    expect(await downloadDriveItem(FILE_URL, out)).toMatchObject({ ok: false, error: { code: ErrorCode.NETWORK_ERROR } });
    expect(existsSync(out)).toBe(false);
  });

  it('discards a rejected Graph token without triggering Teams re-login', async () => {
    mockHttp.mockResolvedValue(err(createError(ErrorCode.AUTH_EXPIRED, 'HTTP 401')));
    expect(await downloadDriveItem(FILE_URL, join(dir, 'x'))).toMatchObject({ ok: false, error: { code: ErrorCode.API_ERROR, retryable: true } });
    expect(invalidateAccessToken).toHaveBeenCalledWith('graph-token');
  });

  it('reports a 403 as access denied, not an auth error', async () => {
    mockHttp.mockResolvedValue(err(createError(ErrorCode.AUTH_REQUIRED, 'HTTP 403')));
    expect(await downloadDriveItem(FILE_URL, join(dir, 'x'))).toMatchObject({ ok: false, error: { code: ErrorCode.ACCESS_DENIED } });
  });

  it('passes through an unavailable Graph token unchanged', async () => {
    const unavailable = err(createError(ErrorCode.AUTH_INTERACTION_REQUIRED, 'Microsoft Graph could not be authorized'));
    vi.mocked(requireGraphTokenAsync).mockResolvedValue(unavailable);
    expect(await downloadDriveItem(FILE_URL, join(dir, 'x'))).toEqual(unavailable);
    expect(mockHttp).not.toHaveBeenCalled();
  });
});
