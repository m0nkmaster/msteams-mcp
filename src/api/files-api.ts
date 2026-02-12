/**
 * Files API client for shared file operations.
 * 
 * Uses the Substrate AllFiles API to retrieve files and links
 * shared in a Teams conversation.
 */

import { httpRequest } from '../utils/http.js';
import { type Result, ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import { requireSubstrateTokenAsync, requireMessageAuth, getTenantId, handleSubstrateError } from '../utils/auth-guards.js';
import { extractObjectId } from '../utils/parsers.js';
import { DEFAULT_FILES_PAGE_SIZE } from '../constants.js';
import { DEFAULT_SUBSTRATE_BASE_URL } from '../utils/api-config.js';
import type { RawAllFilesResponse } from '../types/api-responses.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/** A file shared in a conversation. */
export interface SharedFile {
  /** Item type: "File" or "Link". */
  itemType: 'File' | 'Link';
  /** File name (for files). */
  fileName?: string;
  /** File extension (for files). */
  fileExtension?: string;
  /** URL to access the file or link. For files, this is the SharePoint/OneDrive URL. */
  webUrl?: string;
  /** File size in bytes (for files). */
  fileSize?: number;
  /** Display name of the person who shared the item. */
  sharedBy?: string;
  /** Email of the person who shared the item. */
  sharedByEmail?: string;
  /** When the item was shared (ISO timestamp). */
  sharedTime?: string;
  /** Title (for links). */
  title?: string;
  /** Preview image URL (for images/files with previews). */
  previewUrl?: string;
}

/** Result of getting shared files. */
export interface GetSharedFilesResult {
  conversationId: string;
  files: SharedFile[];
  returned: number;
  /** Continuation token for pagination (if more results available). */
  skipToken?: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// API Functions
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Gets files and links shared in a Teams conversation.
 * 
 * Uses the Substrate AllFiles API which indexes files shared across
 * Teams conversations. Supports both files (SharePoint/OneDrive) and
 * links shared in chat.
 * 
 * @param conversationId - The conversation ID to get files for
 * @param options.pageSize - Number of items per page (default: 25, max: 100)
 * @param options.skipToken - Continuation token for pagination
 */
export async function getSharedFiles(
  conversationId: string,
  options: { pageSize?: number; skipToken?: string } = {}
): Promise<Result<GetSharedFilesResult>> {
  // Need both Substrate token (for the API) and message auth (for user MRI)
  const authResult = requireMessageAuth();
  if (!authResult.ok) return authResult;
  const auth = authResult.value;

  const tokenResult = await requireSubstrateTokenAsync();
  if (!tokenResult.ok) return tokenResult;
  const token = tokenResult.value;

  // Extract user's object ID from their MRI
  const userId = extractObjectId(auth.userMri);
  if (!userId) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'Could not extract user ID from session. Please try logging in again.',
      { suggestions: ['Call teams_login to re-authenticate'] }
    ));
  }

  // Get tenant ID
  const tenantId = getTenantId();
  if (!tenantId) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'Could not determine tenant ID. Please try logging in again.',
      { suggestions: ['Call teams_login to re-authenticate'] }
    ));
  }

  const pageSize = options.pageSize ?? DEFAULT_FILES_PAGE_SIZE;

  // Build the AllFiles API URL
  const userPrincipal = `OID:${userId}@${tenantId}`;
  const params = new URLSearchParams();
  params.set('ThreadId', conversationId);
  params.set('ItemTypes', 'File');
  params.append('ItemTypes', 'Link');
  params.set('PageSize', String(pageSize));

  if (options.skipToken) {
    params.set('$skiptoken', options.skipToken);
  }

  const url = `${DEFAULT_SUBSTRATE_BASE_URL}/AllFiles/api/users('${encodeURIComponent(userPrincipal)}')/AllShared?${params.toString()}`;

  const response = await httpRequest<RawAllFilesResponse>(
    url,
    {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json',
        'Content-Type': 'application/json',
      },
    }
  );

  if (!response.ok) {
    return handleSubstrateError(response);
  }

  const data = response.value.data;
  const rawItems = data.Items;
  const skipToken = data.SkipToken;

  if (!rawItems || rawItems.length === 0) {
    return ok({
      conversationId,
      files: [],
      returned: 0,
      skipToken: undefined,
    });
  }

  const files: SharedFile[] = [];

  for (const item of rawItems) {
    const itemType = item.ItemType;

    if (itemType === 'File') {
      const fileData = item.FileData;
      // FileUrl is the SharePoint URL for uploaded files; WebUrl is often null for chat uploads
      const fileUrl = fileData?.FileUrl ?? fileData?.WebUrl;
      const previewUrl = fileData?.PreviewUrl;
      files.push({
        itemType: 'File',
        fileName: fileData?.FileName,
        fileExtension: fileData?.FileExtension,
        webUrl: fileUrl || undefined,
        fileSize: fileData?.SizeInBytes,
        sharedBy: item.SharedByDisplayName,
        sharedByEmail: item.SharedBySmtp,
        sharedTime: item.SharedDateTime ?? item.SharedTime,
        previewUrl: previewUrl || undefined,
      });
    } else if (itemType === 'Link') {
      const linkData = item.WeblinkData;
      files.push({
        itemType: 'Link',
        webUrl: linkData?.WebUrl,
        title: linkData?.Title,
        sharedBy: item.SharedByDisplayName,
        sharedByEmail: item.SharedBySmtp,
        sharedTime: item.SharedDateTime ?? item.SharedTime,
      });
    }
  }

  return ok({
    conversationId,
    files,
    returned: files.length,
    skipToken: skipToken || undefined,
  });
}
