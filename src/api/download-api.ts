/** Download shared SharePoint/OneDrive files using the existing Teams session. */
import { open, unlink } from 'node:fs/promises';
import { isAbsolute } from 'node:path';
import { createHash } from 'node:crypto';
import { getSharePointToken } from '../auth/token-refresh-http.js';
import { sharePointDownloadUrl } from '../utils/api-config.js';
import { httpRequest } from '../utils/http.js';
import { type Result, ok, err } from '../types/result.js';
import { createError, ErrorCode } from '../types/errors.js';

export interface DownloadFileResult {
  fileName: string;
  outputPath: string;
  size: number;
  contentType: string;
  sha256: string;
}

export async function downloadFile(url: string, outputPath: string): Promise<Result<DownloadFileResult>> {
  if (!isAbsolute(outputPath)) return err(createError(ErrorCode.INVALID_INPUT, 'outputPath must be an absolute file path'));
  let endpoint: ReturnType<typeof sharePointDownloadUrl>;
  try {
    endpoint = sharePointDownloadUrl(url);
  } catch (error) {
    return err(createError(ErrorCode.INVALID_INPUT, error instanceof Error ? error.message : 'Invalid file URL'));
  }
  let token = await getSharePointToken(endpoint.origin);
  if (!token.ok) return token;
  // Exclusive creation protects existing files, including symlink targets.
  let file;
  let complete = false;
  try {
    file = await open(outputPath, 'wx', 0o600);
    const destination = file;
    const request = (bearer: string) => httpRequest<{ size: number; sha256: string }>(endpoint.url, {
      headers: { Authorization: `Bearer ${bearer}` },
      // Do not forward credentials or mistake a sign-in redirect for file contents.
      redirect: 'error',
      // A failed transfer must not replay into a partially written destination.
      maxRetries: 1,
      consumeResponse: async (response, resetTimeout) => {
        const hash = createHash('sha256');
        let size = 0;
        const reader = response.body?.getReader();
        if (reader) {
          try {
            while (true) {
              const next = await reader.read();
              if (next.done) break;
              resetTimeout();
              // Await every write to apply backpressure; memory does not grow with file size.
              await destination.writeFile(next.value);
              hash.update(next.value);
              size += next.value.byteLength;
              resetTimeout();
            }
          } finally {
            await reader.cancel().catch(() => {});
            reader.releaseLock();
          }
        }
        return { size, sha256: hash.digest('hex') };
      },
    });
    let response = await request(token.value);
    // 401 responses never invoke the consumer, so the destination is still empty.
    if (!response.ok && response.error.code === ErrorCode.AUTH_EXPIRED) {
      token = await getSharePointToken(endpoint.origin, true);
      if (!token.ok) return token;
      response = await request(token.value);
    }
    if (!response.ok) return response;
    await file.close();
    complete = true;
    return ok({
      fileName: endpoint.fileName,
      outputPath,
      ...response.value.data,
      contentType: response.value.headers.get('content-type') ?? 'application/octet-stream',
    });
  } catch (error) {
    return err(createError(ErrorCode.INVALID_INPUT, `Could not save file: ${error instanceof Error ? error.message : String(error)}`));
  } finally {
    if (file && !complete) {
      await file.close().catch(() => {});
      await unlink(outputPath).catch(() => {});
    }
  }
}
