/**
 * File downloads via Microsoft Graph.
 *
 * Accepts either form Teams hands out:
 * - Graph drive-item URLs (`https://graph.microsoft.com/v1.0/drives/{driveId}/items/{itemId}`),
 *   used by assignment attachments;
 * - SharePoint/OneDrive web URLs (direct file paths, `Doc.aspx` viewer links,
 *   sharing links), used by shared chat and channel files. These are resolved
 *   through Graph's shares API (`/shares/u!{base64url}/driveItem`).
 *
 * As the Teams Assignments client does, we read the item's
 * `@microsoft.graph.downloadUrl` (a short-lived, pre-authenticated SharePoint
 * link) and stream it to disk.
 *
 * Auth: the Teams client's own Graph token (`Files.ReadWrite.All`), obtained on
 * demand via `requireGraphTokenAsync()`. Like Assignments, it is optional: a
 * failure never surfaces as AUTH_REQUIRED/AUTH_EXPIRED, so it can't trigger the
 * server's auto-login.
 */

import { open, unlink } from 'node:fs/promises';
import { isAbsolute } from 'node:path';
import { createHash } from 'node:crypto';
import { Readable, Transform } from 'node:stream';
import { pipeline } from 'node:stream/promises';
import type { ReadableStream as WebReadableStream } from 'node:stream/web';
import { httpRequest } from '../utils/http.js';
import { type Result, ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import { requireGraphTokenAsync } from '../utils/auth-guards.js';
import { invalidateAccessToken } from '../auth/token-extractor.js';
import { DOWNLOAD_INACTIVITY_TIMEOUT_MS } from '../constants.js';

/** Graph drive items are requested as-is; the Graph token only ever goes to Graph. */
const DRIVE_ITEM_URL = /^https:\/\/graph\.microsoft\.com\/v1\.0\/drives\/[A-Za-z0-9!_-]+\/items\/[A-Za-z0-9!_-]+$/;

/** SharePoint/OneDrive hosts: web URLs to resolve, and the only download-link hosts. */
const SHAREPOINT_HOST = /\.sharepoint(?:-mil)?\.(?:com|us|de|cn)$/;

/**
 * Map an accepted URL to the Graph drive-item URL to request, or null. A
 * SharePoint URL is only ever sent to Graph, base64url-encoded as a share ID.
 */
function toDriveItemUrl(url: string): string | null {
  if (DRIVE_ITEM_URL.test(url)) return url;
  let parsed: URL;
  try { parsed = new URL(url); } catch { return null; }
  if (parsed.protocol !== 'https:' || parsed.username || parsed.password || parsed.port || !SHAREPOINT_HOST.test(parsed.hostname)) {
    return null;
  }
  return `https://graph.microsoft.com/v1.0/shares/u!${Buffer.from(url).toString('base64url')}/driveItem`;
}

export interface DownloadedFile {
  name: string;
  outputPath: string;
  size: number;
  contentType: string;
  sha256: string;
}

interface RawDriveItem {
  name?: string;
  size?: number;
  file?: { mimeType?: string };
  currentUserRole?: { blocksDownload?: boolean };
  '@microsoft.graph.downloadUrl'?: string;
}

/**
 * Downloads a file to `outputPath` (absolute; parent must exist; never
 * overwrites). `url` is a Graph drive-item URL or a SharePoint/OneDrive web
 * URL. Streams to disk, verifies the byte count against Graph's reported size,
 * and removes the partial file on any failure.
 */
export async function downloadFile(url: string, outputPath: string): Promise<Result<DownloadedFile>> {
  const itemUrl = toDriveItemUrl(url);
  if (!itemUrl) {
    return err(createError(ErrorCode.INVALID_INPUT,
      'url must be a Microsoft Graph drive-item URL (https://graph.microsoft.com/v1.0/drives/{id}/items/{id}) or an HTTPS SharePoint/OneDrive file URL (*.sharepoint.com)'));
  }
  if (!isAbsolute(outputPath)) {
    return err(createError(ErrorCode.INVALID_INPUT, 'outputPath must be an absolute file path'));
  }

  const tokenResult = await requireGraphTokenAsync();
  if (!tokenResult.ok) return tokenResult;

  const params = new URLSearchParams({ '$select': 'name,size,file,currentUserRole,content.downloadUrl' });
  const meta = await httpRequest<RawDriveItem>(`${itemUrl}?${params.toString()}`, {
    method: 'GET',
    headers: { Authorization: `Bearer ${tokenResult.value}`, Accept: 'application/json' },
  });
  if (!meta.ok) {
    if (meta.error.code === ErrorCode.AUTH_EXPIRED) {
      invalidateAccessToken(tokenResult.value);
      return err(createError(ErrorCode.API_ERROR,
        'Microsoft Graph rejected its access token; it has been discarded, so a retry will request a new one.',
        { retryable: true }));
    }
    if (meta.error.code === ErrorCode.AUTH_REQUIRED) {
      return err(createError(ErrorCode.ACCESS_DENIED, `You don't have access to this file: ${meta.error.message}`, { retryable: false }));
    }
    if (meta.error.code === ErrorCode.NOT_FOUND) {
      return err(createError(ErrorCode.NOT_FOUND,
        'File not found. It may have been moved or deleted, or the URL is not a file you can access.', { retryable: false }));
    }
    return meta;
  }

  const item = meta.value.data;
  if (!item?.file) {
    return err(createError(ErrorCode.INVALID_INPUT, 'This item is not a file (it may be a folder)', { retryable: false }));
  }
  if (item.currentUserRole?.blocksDownload) {
    return err(createError(ErrorCode.ACCESS_DENIED, 'The owner has blocked downloading this file', { retryable: false }));
  }
  const downloadUrl = item['@microsoft.graph.downloadUrl'];
  let downloadHost = '';
  try { downloadHost = downloadUrl ? new URL(downloadUrl).hostname : ''; } catch { /* handled below */ }
  if (!downloadUrl?.startsWith('https://') || !SHAREPOINT_HOST.test(downloadHost)) {
    return err(createError(ErrorCode.API_ERROR, 'Microsoft Graph did not return a SharePoint download link for this file', { retryable: false }));
  }

  const saved = await streamToFile(downloadUrl, outputPath, item.size);
  if (!saved.ok) return saved;
  return ok({
    name: item.name ?? '',
    outputPath,
    size: saved.value.size,
    contentType: saved.value.contentType ?? item.file.mimeType ?? 'application/octet-stream',
    sha256: saved.value.sha256,
  });
}

/** Stream a pre-authenticated link to a new file, with an inactivity timeout. */
async function streamToFile(
  url: string,
  outputPath: string,
  expectedSize: number | undefined,
): Promise<Result<{ size: number; sha256: string; contentType?: string }>> {
  let file;
  try {
    // Exclusive creation: never overwrite an existing file or follow a symlink.
    file = await open(outputPath, 'wx', 0o600);
  } catch (error) {
    return err(createError(ErrorCode.INVALID_INPUT,
      `Could not create ${outputPath}: ${error instanceof Error ? error.message : String(error)}`,
      { retryable: false, suggestions: ['The parent directory must exist and the file must not already exist'] }));
  }

  const controller = new AbortController();
  let timer = setTimeout(() => controller.abort(), DOWNLOAD_INACTIVITY_TIMEOUT_MS);
  const touch = () => {
    clearTimeout(timer);
    timer = setTimeout(() => controller.abort(), DOWNLOAD_INACTIVITY_TIMEOUT_MS);
  };
  let complete = false;
  try {
    // No Authorization header: the link is pre-authenticated. Refuse redirects
    // so a sign-in page can never be saved as file contents.
    const response = await fetch(url, { redirect: 'error', signal: controller.signal });
    if (!response.ok || !response.body) {
      return err(createError(ErrorCode.API_ERROR, `File download failed: HTTP ${response.status}`,
        { retryable: response.status >= 500 }));
    }
    const hash = createHash('sha256');
    let size = 0;
    const meter = new Transform({
      transform(chunk: Buffer, _encoding, callback) {
        touch();
        hash.update(chunk);
        size += chunk.length;
        callback(null, chunk);
      },
    });
    await pipeline(Readable.fromWeb(response.body as WebReadableStream<Uint8Array>), meter, file.createWriteStream());
    if (expectedSize !== undefined && size !== expectedSize) {
      return err(createError(ErrorCode.API_ERROR,
        `Download was incomplete: received ${size} of ${expectedSize} bytes`, { retryable: true }));
    }
    complete = true;
    return ok({ size, sha256: hash.digest('hex'), contentType: response.headers.get('content-type') ?? undefined });
  } catch (error) {
    const timedOut = controller.signal.aborted;
    return err(createError(timedOut ? ErrorCode.TIMEOUT : ErrorCode.NETWORK_ERROR,
      timedOut
        ? `File download stalled for ${DOWNLOAD_INACTIVITY_TIMEOUT_MS / 1000}s and was cancelled`
        : `File download failed: ${error instanceof Error ? error.message : String(error)}`,
      { retryable: true }));
  } finally {
    clearTimeout(timer);
    await file.close().catch(() => {});
    if (!complete) await unlink(outputPath).catch(() => {});
  }
}
