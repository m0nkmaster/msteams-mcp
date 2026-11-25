/**
 * Chat-message file attachments.
 *
 * Teams stores chat attachments in the sender's OneDrive "Microsoft Teams Chat
 * Files" folder: each file is uploaded via Microsoft Graph, shared with a
 * view-only link granted to the chat's other members, and described to chatsvc
 * as a `properties.files` entry on the outgoing message.
 */

import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { ErrorCode, createError, type McpError } from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';
import { CONFIG_DIR } from '../auth/session-store.js';
import { MAX_SIMPLE_UPLOAD_BYTES, MRI_ORGID_PREFIX, SELF_CHAT_ID } from '../constants.js';
import { httpRequest } from '../utils/http.js';
import { invalidateAccessToken } from '../auth/token-extractor.js';
import { requireGraphTokenAsync, requireMessageAuthWithConfig } from '../utils/auth-guards.js';
import { CHATSVC_API, getSkypeAuthHeaders } from '../utils/api-config.js';
import { extractObjectId } from '../utils/parsers.js';

/** A local file that passed {@link resolveUploadFile} validation. */
export interface ResolvedUploadFile {
  /** Real (symlink-resolved) absolute path to read bytes from. */
  path: string;
  /** File name to use for the OneDrive upload. */
  name: string;
  /** Size in bytes. */
  size: number;
}

/** A file uploaded to OneDrive, as returned by Microsoft Graph. */
export interface UploadedChatFile {
  /** Drive item ID. */
  id: string;
  /** Name Graph stored (may differ from the local name, e.g. "x 2.md"). */
  name: string;
  /** Web URL of the drive item. */
  webUrl: string;
  /** Owning drive ID (from parentReference.driveId). */
  driveId: string;
  /** SharePoint/OneDrive identifiers from the drive item's sharepointIds facet. */
  sharepointIds: {
    listId: string;
    listItemUniqueId: string;
    siteId: string;
    siteUrl: string;
    webId: string;
  };
}

/** The sharing link created for an uploaded file. */
export interface ChatFileShare {
  shareId: string;
  shareUrl: string;
}

/** One entry of a chatsvc message's `properties.files` JSON array. */
export interface ChatsvcFileEntry {
  itemid: string;
  fileName: string;
  fileType: string;
  fileInfo: {
    itemId: string;
    fileUrl: string;
    siteUrl: string;
    serverRelativeUrl: string;
    shareUrl: string;
    shareId: string;
  };
  fileChicletState: { serviceName: string; state: string };
  '@type': string;
  version: number;
  id: string;
  baseUrl: string;
  objectUrl: string;
  openUrl: null;
  type: string;
  title: string;
  state: string;
  altText: null;
  chicletBreadcrumbs: null;
  providerData: string;
  botFileProperties: Record<string, never>;
  isUploadError: null;
  progressComplete: null;
  permissionScope: string;
  filePreview: { previewUrl: string; previewHeight: number; previewWidth: number };
  sharepointIds: {
    driveId: string;
    listId: string;
    listItemUniqueId: string;
    siteId: string;
    siteUrl: string;
    webId: string;
  };
  publication: null;
  site: null;
}

/** Realpath when possible, plain resolve otherwise (protected dirs may not exist). */
function resolveReal(p: string): string {
  try {
    return fs.realpathSync(p);
  } catch {
    return path.resolve(p);
  }
}

/** True when `child` equals `parent` or sits inside it (both absolute). */
function isInside(child: string, parent: string): boolean {
  const rel = path.relative(parent, child);
  return rel === '' || (!rel.startsWith('..') && !path.isAbsolute(rel));
}

export function resolveUploadFile(candidate: string): Result<ResolvedUploadFile> {
  if (!path.isAbsolute(candidate)) {
    return err(createError(ErrorCode.INVALID_INPUT, `Attachment path must be absolute: ${candidate}`));
  }

  const uploadDir = process.env.TEAMS_MCP_UPLOAD_DIR;
  const root = uploadDir || os.homedir();
  if (uploadDir && !path.isAbsolute(uploadDir)) {
    return err(createError(ErrorCode.INVALID_INPUT,
      `TEAMS_MCP_UPLOAD_DIR must be an absolute path: ${uploadDir}`));
  }

  let realRoot: string;
  let realPath: string;
  try {
    realRoot = fs.realpathSync(root);
    realPath = fs.realpathSync(candidate);
  } catch {
    return err(createError(ErrorCode.INVALID_INPUT, `Attachment does not exist: ${candidate}`));
  }

  const relative = path.relative(realRoot, realPath);
  if (relative === '' || relative.startsWith('..') || path.isAbsolute(relative)) {
    return err(createError(ErrorCode.INVALID_INPUT,
      `Attachment must be inside the upload root (${realRoot}): ${candidate}`));
  }

  // Refuse dotfiles and dot-directories (.ssh, .aws, .config, .env, ...): the
  // upload root is broad (a home directory by default), and hidden paths are
  // where credentials live.
  if (relative.split(/[\\/]/).some(segment => segment.startsWith('.'))) {
    return err(createError(ErrorCode.INVALID_INPUT,
      `Attachment path must not contain dotfiles or dot-directories: ${candidate}`));
  }

  // Never upload the server's own config (session state, tokens) or, on
  // Windows, anything under the per-user application-data roots.
  const forbiddenRoots = [CONFIG_DIR];
  if (process.platform === 'win32') {
    if (process.env.APPDATA) forbiddenRoots.push(process.env.APPDATA);
    if (process.env.LOCALAPPDATA) forbiddenRoots.push(process.env.LOCALAPPDATA);
  }
  for (const root of forbiddenRoots) {
    if (isInside(realPath, resolveReal(root))) {
      return err(createError(ErrorCode.INVALID_INPUT,
        `Attachment path is inside a protected directory and cannot be uploaded: ${candidate}`));
    }
  }

  const stat = fs.statSync(realPath);
  if (!stat.isFile()) {
    return err(createError(ErrorCode.INVALID_INPUT, `Attachment is not a regular file: ${candidate}`));
  }
  if (stat.size > MAX_SIMPLE_UPLOAD_BYTES) {
    return err(createError(ErrorCode.INVALID_INPUT,
      `Attachment exceeds the ${MAX_SIMPLE_UPLOAD_BYTES / (1024 * 1024)} MB limit; upload sessions for larger files are not supported yet: ${candidate}`));
  }

  return ok({ path: realPath, name: path.basename(realPath), size: stat.size });
}

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

/**
 * Graph failures must never surface as AUTH_REQUIRED/AUTH_EXPIRED: the server
 * auto-retries those with a full Teams re-login, which cannot help an optional
 * resource like Graph. Mirrors downloadFile in graph-files-api.ts.
 */
function remapGraphError(error: McpError, graphToken: string, action: string): McpError {
  if (error.code === ErrorCode.AUTH_EXPIRED) {
    invalidateAccessToken(graphToken);
    return createError(ErrorCode.API_ERROR,
      `Microsoft Graph rejected its access token while trying to ${action}; it has been discarded, so a retry will request a new one.`,
      { retryable: true });
  }
  if (error.code === ErrorCode.AUTH_REQUIRED) {
    return createError(ErrorCode.ACCESS_DENIED,
      `Microsoft Graph denied permission to ${action}: ${error.message}`,
      { retryable: false });
  }
  return error;
}

interface RawUploadedItem {
  id?: string;
  name?: string;
  webUrl?: string;
  sharepointIds?: {
    listId?: string;
    listItemUniqueId?: string;
    siteId?: string;
    siteUrl?: string;
    webId?: string;
  };
  parentReference?: { driveId?: string };
}

/**
 * Uploads a local file to the sender's OneDrive "Microsoft Teams Chat Files"
 * folder via a simple Graph PUT (files are capped at 4 MB, so no upload
 * session is needed), then reads the item back for the fields chatsvc needs.
 * `conflictBehavior=rename` makes a same-name upload safe, which is why the
 * PUT is never retried: a replay would create a duplicate copy.
 */
export async function uploadChatFile(
  file: ResolvedUploadFile,
  graphToken: string,
): Promise<Result<UploadedChatFile>> {
  const bytes = await fs.promises.readFile(file.path);
  const headers = {
    Authorization: `Bearer ${graphToken}`,
    'Content-Type': 'application/octet-stream',
  };

  const putUrl = `${GRAPH_BASE}/me/drive/root:/Microsoft%20Teams%20Chat%20Files/${encodeURIComponent(file.name)}:/content?@microsoft.graph.conflictBehavior=rename`;
  const put = await httpRequest<{ id?: string }>(putUrl, {
    method: 'PUT',
    headers,
    body: bytes,
    maxRetries: 1,
  });
  if (!put.ok) {
    return err(remapGraphError(put.error, graphToken, `upload "${file.name}"`));
  }
  const itemId = put.value.data.id;
  if (!itemId) {
    return err(createError(ErrorCode.API_ERROR,
      'Microsoft Graph did not return a drive item ID for the uploaded file', { retryable: true }));
  }

  const select = new URLSearchParams({ $select: 'id,name,webUrl,sharepointIds,parentReference' });
  const get = await httpRequest<RawUploadedItem>(
    `${GRAPH_BASE}/me/drive/items/${itemId}?${select.toString()}`,
    { method: 'GET', headers: { Authorization: `Bearer ${graphToken}`, Accept: 'application/json' } },
  );
  if (!get.ok) {
    return err(remapGraphError(get.error, graphToken, `read back "${file.name}"`));
  }

  const item = get.value.data;
  const sp = item.sharepointIds;
  if (!item.name || !item.webUrl || !item.parentReference?.driveId ||
      !sp?.listId || !sp.listItemUniqueId || !sp.siteId || !sp.siteUrl || !sp.webId) {
    return err(createError(ErrorCode.API_ERROR,
      'Microsoft Graph drive item is missing required fields (name, webUrl, sharepointIds, parentReference.driveId)',
      { retryable: false }));
  }

  return ok({
    id: itemId,
    name: item.name,
    webUrl: item.webUrl,
    driveId: item.parentReference.driveId,
    sharepointIds: {
      listId: sp.listId,
      listItemUniqueId: sp.listItemUniqueId,
      siteId: sp.siteId,
      siteUrl: sp.siteUrl,
      webId: sp.webId,
    },
  });
}

interface RawPermission {
  id?: string;
  link?: { webUrl?: string };
}

/**
 * Creates the view-only sharing link Teams attaches to chat files, then
 * grants read to the chat's other members. With no recipients (self-notes)
 * the link is still created — matching what the Teams client does.
 */
export async function shareChatFile(
  itemId: string,
  recipientObjectIds: string[],
  graphToken: string,
): Promise<Result<ChatFileShare>> {
  const headers = {
    Authorization: `Bearer ${graphToken}`,
    'Content-Type': 'application/json',
    Accept: 'application/json',
  };

  const link = await httpRequest<RawPermission>(
    `${GRAPH_BASE}/me/drive/items/${itemId}/createLink`,
    {
      method: 'POST',
      headers,
      body: JSON.stringify({ type: 'view', scope: 'users' }),
      maxRetries: 1,
    },
  );
  if (!link.ok) {
    return err(remapGraphError(link.error, graphToken, 'share the uploaded file'));
  }
  const permission = link.value.data;
  const shareUrl = permission.link?.webUrl;
  if (!permission.id || !shareUrl) {
    return err(createError(ErrorCode.API_ERROR,
      'Microsoft Graph did not return a sharing link for the uploaded file', { retryable: true }));
  }

  if (recipientObjectIds.length > 0) {
    const shareId = `u!${Buffer.from(shareUrl).toString('base64url')}`;
    const grant = await httpRequest<unknown>(
      `${GRAPH_BASE}/shares/${shareId}/permission/grant`,
      {
        method: 'POST',
        headers,
        body: JSON.stringify({
          recipients: recipientObjectIds.map(objectId => ({ objectId })),
          roles: ['read'],
        }),
        maxRetries: 1,
      },
    );
    if (!grant.ok) {
      return err(remapGraphError(grant.error, graphToken, 'share the uploaded file with the chat'));
    }
  }

  return ok({ shareId: permission.id, shareUrl });
}

/**
 * Maps an uploaded+shared file to the chatsvc `properties.files` entry, in the
 * exact shape and key order the Teams client emits (see the
 * chatsvcFilesPropertyEntry fixture). `itemid`/`id` are the SharePoint list
 * item unique ID; `fileInfo.itemId` is the Graph drive item ID.
 */
export function buildFilesPropertyEntry(
  file: UploadedChatFile,
  share: ChatFileShare,
): ChatsvcFileEntry {
  const sp = file.sharepointIds;
  const fileType = path.extname(file.name).slice(1).toLowerCase();
  const itemid = sp.listItemUniqueId;
  const baseUrl = sp.siteUrl.replace(/\/+$/, '') + '/';
  const fileUrl = file.webUrl;

  return {
    itemid,
    fileName: file.name,
    fileType,
    fileInfo: {
      itemId: file.id,
      fileUrl,
      siteUrl: baseUrl,
      serverRelativeUrl: '',
      shareUrl: share.shareUrl,
      shareId: share.shareId,
    },
    fileChicletState: { serviceName: 'p2p', state: 'active' },
    '@type': 'http://schema.skype.com/File',
    version: 2,
    id: itemid,
    baseUrl,
    objectUrl: fileUrl,
    openUrl: null,
    type: fileType,
    title: file.name,
    state: 'active',
    altText: null,
    chicletBreadcrumbs: null,
    providerData: '',
    botFileProperties: {},
    isUploadError: null,
    progressComplete: null,
    permissionScope: 'SpecificPeople',
    filePreview: { previewUrl: '', previewHeight: 0, previewWidth: 0 },
    sharepointIds: {
      driveId: file.driveId,
      listId: sp.listId,
      listItemUniqueId: sp.listItemUniqueId,
      siteId: sp.siteId,
      siteUrl: sp.siteUrl,
      webId: sp.webId,
    },
    publication: null,
    site: null,
  };
}

interface RawThreadDetails {
  members?: Array<{ id?: string }>;
}

/**
 * Object IDs of a chat's other members, for Graph sharing grants. Reads the
 * chatsvc thread endpoint (the conversations list carries no members). Only
 * `8:orgid:` user MRIs map to Graph object IDs; bots, tags and the caller are
 * excluded. Errors pass through unchanged — this is core Teams auth, so
 * AUTH_REQUIRED/AUTH_EXPIRED are legitimate here.
 */
export async function getChatMemberObjectIds(
  conversationId: string,
): Promise<Result<string[]>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, region, baseUrl } = authResult.value;

  const response = await httpRequest<RawThreadDetails>(
    `${CHATSVC_API.thread(region, conversationId, baseUrl)}?view=msnp24Equivalent`,
    {
      method: 'GET',
      headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken, baseUrl),
    },
  );
  if (!response.ok) {
    return response;
  }

  const members = response.value.data.members;
  if (!Array.isArray(members)) {
    return err(createError(ErrorCode.API_ERROR,
      'Teams did not return a members list for this chat', { retryable: true }));
  }

  const objectIds = new Set<string>();
  for (const member of members) {
    const id = member?.id;
    if (typeof id !== 'string' || !id.startsWith(MRI_ORGID_PREFIX) || id === auth.userMri) {
      continue;
    }
    const objectId = extractObjectId(id);
    if (objectId) {
      objectIds.add(objectId);
    }
  }
  return ok([...objectIds]);
}

/**
 * Validates every path, then uploads, shares and describes each file for a
 * chatsvc `properties.files` entry. All local validation and the channel
 * check happen before any network call; files are processed sequentially.
 * A mid-way failure returns an error naming the file — the message is never
 * sent, but earlier uploads stay in OneDrive "Microsoft Teams Chat Files".
 */
export async function prepareChatAttachments(
  conversationId: string,
  paths: string[],
): Promise<Result<ChatsvcFileEntry[]>> {
  if (paths.length === 0) {
    return ok([]);
  }

  const files: ResolvedUploadFile[] = [];
  for (const candidate of paths) {
    const resolved = resolveUploadFile(candidate);
    if (!resolved.ok) {
      return err(resolved.error);
    }
    files.push(resolved.value);
  }

  if (conversationId.endsWith('@thread.tacv2') || conversationId.endsWith('@thread.skype')) {
    return err(createError(ErrorCode.INVALID_INPUT,
      'Attachments are supported in 1:1, group and meeting chats and self-notes; channel files live in the team\'s SharePoint'));
  }

  // Optional resource: requireGraphTokenAsync never returns AUTH_* (that would
  // trigger the server's Teams re-login), same contract as downloadFile.
  const tokenResult = await requireGraphTokenAsync();
  if (!tokenResult.ok) {
    return tokenResult;
  }
  const graphToken = tokenResult.value;

  // Self-notes have no other members to share with; everyone else gets a
  // read grant for the chat's members.
  let recipients: string[] = [];
  if (conversationId !== SELF_CHAT_ID) {
    const members = await getChatMemberObjectIds(conversationId);
    if (!members.ok) {
      return err(members.error);
    }
    recipients = members.value;
  }

  const entries: ChatsvcFileEntry[] = [];
  for (const file of files) {
    const uploaded = await uploadChatFile(file, graphToken);
    if (!uploaded.ok) {
      return err(attachmentFailure(file, uploaded.error));
    }
    const share = await shareChatFile(uploaded.value.id, recipients, graphToken);
    if (!share.ok) {
      return err(attachmentFailure(file, share.error));
    }
    entries.push(buildFilesPropertyEntry(uploaded.value, share.value));
  }
  return ok(entries);
}

/** Wraps a per-file Graph failure: names the file and states nothing was sent. */
function attachmentFailure(file: ResolvedUploadFile, cause: McpError): McpError {
  return createError(cause.code,
    `Failed to attach "${file.name}": ${cause.message} Nothing was sent — the message was not delivered. Files already uploaded remain in your OneDrive "Microsoft Teams Chat Files" folder.`,
    { retryable: cause.retryable, retryAfterMs: cause.retryAfterMs, suggestions: cause.suggestions });
}
