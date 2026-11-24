import { z } from 'zod';
import { CSA_API, getCsaHeaders } from '../utils/api-config.js';
import { getApiConfig, requireCsaAuth } from '../utils/auth-guards.js';
import { httpRequest } from '../utils/http.js';
import { createError, ErrorCode } from '../types/errors.js';
import { err, ok, type Result } from '../types/result.js';
import { getConversationProperties, extractParticipantNames } from './chatsvc-api.js';

const itemSchema = z.object({
  conversationId: z.string().min(1).optional(),
  createdTime: z.number().optional().catch(undefined),
  lastUpdatedTime: z.number().optional().catch(undefined),
});
const folderSchema = z.object({
  id: z.string().min(1),
  name: z.string().nullish(),
  folderType: z.string(),
  isDeleted: z.boolean().optional(),
  conversationFolderItems: z.array(itemSchema.catch({})).nullish(),
});
const hierarchySchema = z.object({
  folderHierarchyVersion: z.number().optional(),
  conversationFolders: z.array(folderSchema).transform(folders => folders.filter(folder => !folder.isDeleted)),
});
type Folder = z.infer<typeof folderSchema>;
type Item = z.infer<typeof itemSchema> & { conversationId: string; displayName?: string; conversationType?: string };

async function folderRequest(method: 'GET' | 'POST' | 'DELETE', body?: object, folderId?: string) {
  const credentials = requireCsaAuth();
  if (!credentials.ok) return credentials;
  const { auth, csaToken } = credentials.value;
  const { csaRegion, baseUrl } = getApiConfig();
  return httpRequest<unknown>(CSA_API.conversationFolders(csaRegion, baseUrl, folderId), {
    method, headers: getCsaHeaders(auth.skypeToken, csaToken, baseUrl),
    ...(body && { body: JSON.stringify(body), maxRetries: 1 }),
  });
}

export async function getFolderHierarchy() {
  const response = await folderRequest('GET');
  if (!response.ok) return response;
  const parsed = hierarchySchema.safeParse(response.value.data);
  return parsed.success ? ok(parsed.data) : err(createError(ErrorCode.API_ERROR, 'Unexpected CSA section hierarchy.'));
}

function summary(folder: Folder) {
  return { id: folder.id, name: folder.name ?? null, type: folder.folderType,
    itemCount: isEditable(folder) ? folder.conversationFolderItems?.length ?? null : null };
}

function resolveSection(folders: Folder[], selector?: string): Result<Folder> {
  if (selector === undefined) {
    const favorites = folders.find(f => f.folderType === 'Favorites');
    return favorites ? ok(favorites) : err(createError(ErrorCode.NOT_FOUND, 'Could not find Favorites folder'));
  }
  const idMatch = folders.find(f => f.id === selector);
  if (idMatch) return ok(idMatch);
  const matches = folders.filter(f => f.name?.toLowerCase() === selector.toLowerCase());
  if (matches.length === 1) return ok(matches[0]);
  if (!matches.length) return err(createError(ErrorCode.NOT_FOUND, 'Section not found: ' + selector));
  return err(createError(ErrorCode.INVALID_INPUT, 'Ambiguous section name. Use one of these IDs: ' + matches.map(f => f.id).join(', ')));
}

function members(folder: Folder): Result<Item[]> {
  if (!isEditable(folder)) return err(createError(ErrorCode.INVALID_INPUT,
    'Membership is unavailable for this system section. Choose Favorites or a custom section.'));
  const items: Item[] = [];
  if (!folder.conversationFolderItems) return err(createError(ErrorCode.API_ERROR, 'Teams did not return section membership.'));
  for (const item of folder.conversationFolderItems) {
    if (!item.conversationId) continue;
    items.push({ ...item, conversationId: item.conversationId });
  }
  return ok(items);
}

export async function getSectionMembers(selector: string) {
  const state = await getFolderHierarchy();
  if (!state.ok) return state;
  const folder = resolveSection(state.value.conversationFolders, selector);
  if (!folder.ok) return folder;
  const items = members(folder.value);
  if (!items.ok) return items;
  const section = summary(folder.value);
  const skippedItems = (section.itemCount ?? 0) - items.value.length;
  return ok({ section, items: items.value, skippedItems });
}

export async function enrichSectionItems(items: Item[]) {
  await Promise.allSettled(items.map(async item => {
    const props = await getConversationProperties(item.conversationId);
    if (props.ok) {
      item.displayName = props.value.displayName;
      item.conversationType = props.value.conversationType;
    }
    if (!item.displayName) {
      const names = await extractParticipantNames(item.conversationId);
      if (names.ok && names.value) item.displayName = names.value;
    }
  }));
  return items;
}

export async function getSections(selector?: string): Promise<Result<Record<string, unknown>>> {
  if (selector) {
    const result = await getSectionMembers(selector);
    if (!result.ok) return result;
    return ok({ ...result.value, items: await enrichSectionItems(result.value.items) });
  }
  const state = await getFolderHierarchy();
  return state.ok ? ok({ sections: state.value.conversationFolders.map(summary) }) : state;
}

export type SectionUpdate =
  | { action: 'create'; name: string }
  | { action: 'delete'; section: string }
  | { action: 'add' | 'remove'; section?: string; conversationId: string };

export async function updateSection(input: SectionUpdate): Promise<Result<Record<string, unknown>>> {
  const state = await getFolderHierarchy();
  if (!state.ok) return state;
  const { folderHierarchyVersion, conversationFolders } = state.value;
  if (folderHierarchyVersion === undefined) return err(createError(ErrorCode.API_ERROR, 'Teams omitted the folder hierarchy version.'));
  const base = { folderHierarchyVersion, supportsAdditionalSystemGeneratedFolders: true, supportsSliceItems: true };
  if (input.action === 'create') {
    if (conversationFolders.some(f => f.name?.toLowerCase() === input.name.toLowerCase())) {
      return err(createError(ErrorCode.INVALID_INPUT, 'A section with this name already exists.'));
    }
    const response = await folderRequest('POST', { ...base, folderName: input.name,
      sortType: 'UserDefinedCustomOrder', isExpanded: true, actions: [{ action: 'AddFolder' }] });
    if (!response.ok) return response;
    const created = hierarchySchema.safeParse(response.value.data);
    const folder = created.success && created.data.conversationFolders.find(f =>
      f.name === input.name && f.folderType === 'UserCreated' && !conversationFolders.some(old => old.id === f.id));
    if (!folder) return err(createError(ErrorCode.API_ERROR, 'Could not confirm the created section. List sections before retrying.'));
    return ok({ section: summary(folder) });
  }

  const resolved = resolveSection(conversationFolders, input.section);
  if (!resolved.ok) return resolved;
  const folder = resolved.value;
  if (input.action === 'delete') {
    if (input.section !== folder.id || folder.folderType !== 'UserCreated') {
      return err(createError(ErrorCode.INVALID_INPUT, 'Deletion requires the ID of a user-created section.'));
    }
    const response = await folderRequest('DELETE', base, folder.id);
    return response.ok ? ok({ deletedSectionId: folder.id }) : response;
  }
  if (!isEditable(folder)) return err(createError(ErrorCode.INVALID_INPUT, 'Membership is editable only in Favorites and user-created sections.'));
  const items = members(folder);
  if (!items.ok) return items;
  const { conversationId } = input;
  const adding = input.action === 'add';
  const result = { sectionId: folder.id, conversationId, action: input.action };
  if (items.value.some(item => item.conversationId === conversationId) === adding) return ok({ ...result, changed: false });
  const actions = [{ action: adding ? 'AddItem' : 'RemoveItem', folderId: folder.id, itemId: conversationId }];
  if (adding) {
    for (const source of conversationFolders.filter(f => f.id !== folder.id && isEditable(f))) {
      const sourceItems = members(source);
      if (!sourceItems.ok) return sourceItems;
      if (sourceItems.value.some(item => item.conversationId === conversationId)) {
        actions.push({ action: 'RemoveItem', folderId: source.id, itemId: conversationId });
      }
    }
  }
  const response = await folderRequest('POST', { ...base, actions });
  return response.ok ? ok({ ...result, changed: true }) : response;
}

function isEditable(folder: Folder) {
  return folder.folderType === 'Favorites' || folder.folderType === 'UserCreated';
}
