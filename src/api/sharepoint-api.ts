/**
 * SharePoint/OneDrive file upload API via Microsoft Graph.
 *
 * Teams file attachments work in two steps: (1) upload the file to the user's
 * OneDrive "Microsoft Teams Chat Files" folder via the Graph API, then (2) send
 * a chat message with a `files` property referencing the uploaded file.
 *
 * We use the Graph API (`graph.microsoft.com`) rather than the SharePoint REST
 * API directly because:
 * - The Graph token is already in the MSAL cache (broad delegated permissions)
 * - No need to construct tenant-specific SharePoint URLs manually
 * - The Graph response includes all SharePoint URLs needed for the `files` property
 *
 * Reverse-engineered from Teams web client network interception (2026-07-28).
 */

import { readFile } from 'node:fs/promises';
import { basename, extname } from 'node:path';
import { httpRequest } from '../utils/http.js';
import { ErrorCode, createError } from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';
import { getValidGraphToken } from '../auth/token-extractor.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/** Subset of the Graph API DriveItem response that we need. */
export interface DriveItem {
  id: string;
  name: string;
  size?: number;
  webUrl?: string;
  sharepointIds?: {
    siteId?: string;
    siteUrl?: string;
    webId?: string;
    listId?: string;
    listItemUniqueId?: string;
  };
  parentReference?: {
    driveId?: string;
    sharepointIds?: {
      siteId?: string;
      siteUrl?: string;
      webId?: string;
      listId?: string;
      listItemUniqueId?: string;
    };
  };
}

/** Result of uploading a file. */
export interface UploadFileResult {
  /** The DriveItem ID from Graph/SharePoint. */
  itemId: string;
  /** The file name (may differ from input if renamed due to conflict). */
  fileName: string;
  /** File extension without the dot (e.g., "pdf"). */
  fileType: string;
  /** File size in bytes. */
  fileSize?: number;
  /** The SharePoint personal site base URL (e.g., "https://tenant-my.sharepoint.com/personal/user_domain_com/"). */
  baseUrl: string;
  /** The full SharePoint URL to the file. */
  objectUrl: string;
  /** The web URL from Graph (may differ slightly from objectUrl). */
  webUrl?: string;
  /** SharePoint list item unique ID (used in the `files` property). */
  listItemUniqueId?: string;
  /** The JSON-encoded `files` property string for the chatsvc message body. */
  filesProperty: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────────────────────────────────────

/** Graph API base URL. */
const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

/** The OneDrive folder Teams uses for chat file attachments. */
const TEAMS_CHAT_FILES_FOLDER = 'Microsoft Teams Chat Files';

/** Maximum file size for simple upload (4 MB — Graph API limit for single PUT). */
const MAX_SIMPLE_UPLOAD_SIZE = 4 * 1024 * 1024;

// ─────────────────────────────────────────────────────────────────────────────
// Helper Functions
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Gets the file extension without the leading dot (e.g., "pdf" from "doc.pdf").
 */
function getFileExtension(fileName: string): string {
  const ext = extname(fileName).toLowerCase().replace(/^\./, '');
  return ext || 'file';
}

/**
 * Extracts the SharePoint personal site base URL from a DriveItem response.
 *
 * The base URL looks like: `https://{tenant}-my.sharepoint.com/personal/{user_folder}/`
 */
function extractBaseUrl(driveItem: DriveItem): string | null {
  // Try sharepointIds.siteUrl first
  const siteUrl = driveItem.sharepointIds?.siteUrl
    ?? driveItem.parentReference?.sharepointIds?.siteUrl;
  if (siteUrl) {
    return siteUrl.endsWith('/') ? siteUrl : `${siteUrl}/`;
  }

  // Fallback: parse from webUrl
  if (driveItem.webUrl) {
    try {
      const url = new URL(driveItem.webUrl);
      // webUrl is like: https://{tenant}-my.sharepoint.com/personal/{user}/Documents/Microsoft Teams Chat Files/{file}
      const pathParts = url.pathname.split('/');
      // Find "personal" in the path and take the next segment
      const personalIndex = pathParts.indexOf('personal');
      if (personalIndex >= 0 && pathParts[personalIndex + 1]) {
        const base = `${url.protocol}//${url.host}/personal/${pathParts[personalIndex + 1]}/`;
        return base;
      }
    } catch {
      // Ignore parse errors
    }
  }

  return null;
}

/**
 * Builds the `files` property JSON string from a Graph API DriveItem response.
 *
 * This is the exact format the Teams chatsvc API expects in the message body's
 * `properties.files` field — a JSON-encoded string (not an array).
 */
export function buildFilesProperty(driveItem: DriveItem): string {
  const baseUrl = extractBaseUrl(driveItem) ?? '';
  const fileName = driveItem.name;
  const fileType = getFileExtension(fileName);
  const itemId = driveItem.id;
  const listItemUniqueId = driveItem.sharepointIds?.listItemUniqueId
    ?? driveItem.parentReference?.sharepointIds?.listItemUniqueId
    ?? itemId;

  const objectUrl = driveItem.webUrl
    ?? `${baseUrl}Documents/${TEAMS_CHAT_FILES_FOLDER}/${encodeURIComponent(fileName)}`;

  const file = {
    itemid: itemId,
    fileName,
    fileType,
    fileInfo: {
      itemId: null,
      fileUrl: objectUrl,
      siteUrl: baseUrl,
      serverRelativeUrl: '',
      shareUrl: null,
      shareId: null,
    },
    fileChicletState: {
      serviceName: 'p2p',
      state: 'active',
    },
    '@type': 'http://schema.skype.com/File',
    version: 2,
    id: itemId,
    baseUrl,
    objectUrl,
    type: fileType,
    title: fileName,
    state: 'active',
    chicletBreadcrumbs: null,
    providerData: '',
    botFileProperties: {},
    isUploadError: null,
    progressComplete: null,
    permissionScope: 'anonymous',
    filePreview: {
      previewUrl: '',
      previewHeight: 0,
      previewWidth: 0,
    },
    sharepointIds: {
      listId: driveItem.sharepointIds?.listId ?? null,
      listItemUniqueId,
      siteId: driveItem.sharepointIds?.siteId ?? null,
      siteUrl: driveItem.sharepointIds?.siteUrl ?? null,
      webId: driveItem.sharepointIds?.webId ?? null,
    },
    publication: null,
    site: null,
  };

  return JSON.stringify([file]);
}

// ─────────────────────────────────────────────────────────────────────────────
// File Upload
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Uploads a local file to the user's OneDrive "Microsoft Teams Chat Files" folder.
 *
 * Uses the Graph API simple upload (PUT with content). For files larger than 4 MB,
 * a session-based upload would be needed (not yet implemented).
 *
 * @param filePath - Absolute or relative path to the local file
 * @returns Upload result with item metadata and the `files` property string
 */
export async function uploadFile(filePath: string): Promise<Result<UploadFileResult>> {
  // Validate auth
  const graphToken = getValidGraphToken();
  if (!graphToken) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'ACTION REQUIRED: No valid Microsoft Graph token. You MUST call teams_login to authenticate before uploading files.',
      { suggestions: ['Call teams_login to authenticate via browser'] }
    ));
  }

  // Read the file
  let fileBuffer: Buffer;
  try {
    fileBuffer = await readFile(filePath);
  } catch (error) {
    return err(createError(
      ErrorCode.INVALID_INPUT,
      `Failed to read file "${filePath}": ${error instanceof Error ? error.message : String(error)}`,
      { retryable: false }
    ));
  }

  // Check file size (Graph API simple upload limit is 4 MB)
  if (fileBuffer.length > MAX_SIMPLE_UPLOAD_SIZE) {
    return err(createError(
      ErrorCode.INVALID_INPUT,
      `File "${filePath}" is ${Math.round(fileBuffer.length / 1024 / 1024)} MB. Maximum size for upload is ${MAX_SIMPLE_UPLOAD_SIZE / 1024 / 1024} MB. Large file upload is not yet supported.`,
      { retryable: false }
    ));
  }

  const fileName = basename(filePath);

  // Build the Graph API upload URL
  // PUT /me/drive/root:/Microsoft Teams Chat Files/{filename}:/content
  const uploadUrl = `${GRAPH_BASE_URL}/me/drive/root:/${encodeURIComponent(TEAMS_CHAT_FILES_FOLDER)}/${encodeURIComponent(fileName)}:/content`;

  const response = await httpRequest<DriveItem>(uploadUrl, {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${graphToken}`,
      'Content-Type': 'application/octet-stream',
    },
    body: new Uint8Array(fileBuffer),
    maxRetries: 1, // Don't retry uploads — could result in duplicate files
  });

  if (!response.ok) {
    return response;
  }

  const driveItem = response.value.data;
  const baseUrl = extractBaseUrl(driveItem);
  if (!baseUrl) {
    return err(createError(
      ErrorCode.UNKNOWN,
      `File uploaded successfully but could not determine SharePoint base URL from the response. Item ID: ${driveItem.id}`,
      { retryable: false }
    ));
  }

  const filesProperty = buildFilesProperty(driveItem);

  return ok({
    itemId: driveItem.id,
    fileName: driveItem.name,
    fileType: getFileExtension(driveItem.name),
    fileSize: driveItem.size,
    baseUrl,
    objectUrl: driveItem.webUrl ?? `${baseUrl}Documents/${TEAMS_CHAT_FILES_FOLDER}/${encodeURIComponent(driveItem.name)}`,
    webUrl: driveItem.webUrl,
    listItemUniqueId: driveItem.sharepointIds?.listItemUniqueId
      ?? driveItem.parentReference?.sharepointIds?.listItemUniqueId
      ?? driveItem.id,
    filesProperty,
  });
}

/**
 * Uploads multiple files and returns their combined `files` property string.
 *
 * The `files` property in the chatsvc message body is a JSON-encoded array.
 * When multiple files are attached, their entries are merged into a single array.
 *
 * @param filePaths - Array of local file paths
 * @returns Combined `files` property string and per-file upload results
 */
export async function uploadFiles(
  filePaths: string[]
): Promise<Result<{ filesProperty: string; uploads: UploadFileResult[] }>> {
  const uploads: UploadFileResult[] = [];
  const fileEntries: unknown[] = [];

  for (const filePath of filePaths) {
    const result = await uploadFile(filePath);
    if (!result.ok) {
      return result;
    }
    uploads.push(result.value);

    // Parse the filesProperty (which is JSON.stringify([singleFile])) and merge
    try {
      const parsed = JSON.parse(result.value.filesProperty) as unknown[];
      fileEntries.push(...parsed);
    } catch {
      return err(createError(
        ErrorCode.UNKNOWN,
        `Failed to parse files property for uploaded file "${result.value.fileName}"`,
        { retryable: false }
      ));
    }
  }

  return ok({
    filesProperty: JSON.stringify(fileEntries),
    uploads,
  });
}
