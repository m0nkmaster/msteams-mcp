import { beforeEach, expect, it, vi } from 'vitest';
import { invokeTool } from './registry.js';
import { httpRequest } from '../utils/http.js';
import { ok, err } from '../types/result.js';
import { createError, ErrorCode } from '../types/errors.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({
  requireCsaAuth: () => ok({ auth: { skypeToken: 'test' }, csaToken: 'test' }),
  getApiConfig: () => ({ csaRegion: 'amer', baseUrl: 'https://teams.example' }),
  requireSubstrateTokenAsync: () => ok('test'),
  getTenantId: () => undefined,
  getTeamsBaseUrl: () => undefined,
}));
vi.mock('../api/chatsvc-api.js', async importOriginal => ({
  ...await importOriginal<typeof import('../api/chatsvc-api.js')>(),
  getConversationProperties: () => ok({ displayName: 'General' }),
}));

const http = vi.mocked(httpRequest);
const response = (data: unknown) => ok({ status: 200, headers: new Headers(), data });
const folder = (id: string, name: string, ids: string[], folderType = 'UserCreated') => ({
  id, name, folderType, conversationFolderItems: ids.map(conversationId => ({ conversationId })),
});
const projects = folder('p/1', 'Projects', ['channel-a']);
const favorites = folder('fav', 'Favorites', ['chat-b'], 'Favorites');
function hierarchy(...conversationFolders: unknown[]) {
  http.mockResolvedValueOnce(response({ folderHierarchyVersion: 42, conversationFolders }));
}
const update = (input: object) => invokeTool('teams_update_section', input);
const search = (input: object = {}) => invokeTool('teams_search', { section: 'Projects', query: 'launch', ...input });
const body = (index: number) => JSON.parse(String(http.mock.calls[index][1]?.body));
beforeEach(() => http.mockReset());

it('lists sections without fetching names and inspects one by case-insensitive name', async () => {
  hierarchy(projects, favorites, { ...projects, id: 'old', isDeleted: true });
  expect(await invokeTool('teams_get_sections', {})).toMatchObject({ data: { sections: [
    { id: 'p/1', itemCount: 1 }, { id: 'fav', type: 'Favorites' },
  ] } });
  hierarchy(projects);
  expect(await invokeTool('teams_get_sections', { section: 'projects' })).toMatchObject({
    data: { items: [{ conversationId: 'channel-a', displayName: 'General' }] },
  });
});

it('rejects ambiguous names, protected deletion, deletion by name and invalid action inputs', async () => {
  hierarchy(projects, folder('other', 'Projects', []));
  expect(await search()).toMatchObject({ success: false, error: { message: expect.stringContaining('other') } });
  for (const section of ['fav', 'Projects']) {
    hierarchy(projects, favorites);
    expect(await update({ action: 'delete', section })).toMatchObject({ success: false });
  }
  expect(await update({ action: 'add', section: 'Projects' })).toMatchObject({ success: false });
  expect(http).toHaveBeenCalledTimes(3);
});

it('creates and deletes a section using its returned ID and a versioned request', async () => {
  hierarchy(favorites);
  hierarchy(projects);
  expect(await update({ action: 'create', name: 'Projects' })).toMatchObject({ data: { section: { id: 'p/1' } } });
  expect(body(1)).toMatchObject({ folderHierarchyVersion: 42, folderName: 'Projects', actions: [{ action: 'AddFolder' }] });
  hierarchy(projects);
  http.mockResolvedValueOnce(response({}));
  expect(await update({ action: 'delete', section: 'p/1' })).toMatchObject({ data: { deletedSectionId: 'p/1' } });
  expect(http.mock.calls[3]).toMatchObject([expect.stringContaining('/p%2F1?'), { method: 'DELETE', maxRetries: 1 }]);
});

it('moves with one atomic request and removes only the requested membership', async () => {
  hierarchy(projects, favorites);
  http.mockResolvedValueOnce(response({}));
  expect(await update({ action: 'add', section: 'Projects', conversationId: 'chat-b' })).toMatchObject({ data: { changed: true } });
  expect(body(1).actions).toEqual([
    { action: 'AddItem', folderId: 'p/1', itemId: 'chat-b' },
    { action: 'RemoveItem', folderId: 'fav', itemId: 'chat-b' },
  ]);
  hierarchy(projects);
  http.mockResolvedValueOnce(response({}));
  await update({ action: 'remove', section: 'p/1', conversationId: 'channel-a' });
  expect(body(3).actions).toEqual([{ action: 'RemoveItem', folderId: 'p/1', itemId: 'channel-a' }]);
  hierarchy(projects);
  expect(await update({ action: 'add', section: 'Projects', conversationId: 'channel-a' })).toMatchObject({ data: { changed: false } });
  hierarchy(projects);
  expect(await update({ action: 'remove', section: 'Projects', conversationId: 'missing' })).toMatchObject({ data: { changed: false } });
  expect(http).toHaveBeenCalledTimes(6);
});

it('does not replay a create after a timeout or create a duplicate on a later call', async () => {
  hierarchy(favorites);
  http.mockResolvedValueOnce(err(createError(ErrorCode.TIMEOUT, 'timeout')));
  expect(await update({ action: 'create', name: 'Projects' })).toMatchObject({ success: false });
  hierarchy(projects);
  expect(await update({ action: 'create', name: 'Projects' })).toMatchObject({ success: false });
  expect(http.mock.calls[1][1]?.maxRetries).toBe(1);
  expect(http).toHaveBeenCalledTimes(3);
});

function hit(channel: string, id: string) {
  return { Id: id, Summary: 'launch summary', Source: {
    ChannelName: 'General', ClientThreadId: 'thread-' + id, ClientConversationId: channel + ';messageid=123',
  } };
}
function page(Results: unknown[], Total?: number) {
  http.mockResolvedValueOnce(response({ EntitySets: [{ ResultSets: [{ Total, Results }] }] }));
}

it('matches the parent channel across pages and resumes mid-page after rejected hits', async () => {
  hierarchy(projects);
  page([hit('channel-b', 'wrong'), { Summary: '' }], 5);
  page([hit('channel-a', 'a'), hit('channel-b', 'wrong-again'), hit('channel-a', 'b')], 5);
  expect(await search({ size: 1 })).toMatchObject({
    data: { results: [{ id: 'a' }], pagination: { hasMore: true, nextFrom: 3, returned: 1 } },
  });
  expect(body(2).entityRequests[0].from).toBe(2);
  hierarchy(projects);
  page([hit('channel-b', 'wrong-again'), hit('channel-a', 'b')], 5);
  expect(await search({ size: 1, from: 3 })).toMatchObject({ data: { results: [{ id: 'b' }], pagination: { hasMore: false } } });
});

it('recognizes a channel reply identified by SkypeGroupId when ClientConversationId is absent', async () => {
  hierarchy(projects);
  page([{ Id: 'reply', Summary: 'launch summary', Source: {
    ClientThreadId: 'reply-thread',
    Extensions: { SkypeSpaces_ConversationPost_Extension_SkypeGroupId: 'channel-a' },
  } }], 1);
  expect(await search()).toMatchObject({ data: { results: [{ id: 'reply' }], pagination: { hasMore: false } } });
});

it('returns a usable continuation after a bounded scan with zero matches', async () => {
  hierarchy(projects);
  for (let i = 0; i < 3; i++) page([hit('channel-b', String(i))], 20);
  expect(await search()).toMatchObject({ data: { results: [], pagination: { nextFrom: 3, hasMore: true } } });
});

it('continues past short pages without totals until the upstream page is empty', async () => {
  hierarchy(projects);
  page([hit('channel-b', 'wrong')]);
  page([hit('channel-a', 'found')]);
  page([]);
  expect(await search()).toMatchObject({
    data: { results: [{ id: 'found' }], pagination: { hasMore: false } },
  });
  expect(body(2).entityRequests[0].from).toBe(1);
  expect(body(3).entityRequests[0].from).toBe(2);
});

it('distinguishes unreadable membership from partial section search results', async () => {
  hierarchy({ ...projects, conversationFolderItems: [{}] });
  expect(await search()).toMatchObject({ success: false, error: { code: ErrorCode.API_ERROR } });
  expect(http).toHaveBeenCalledTimes(1);

  hierarchy({ ...projects, conversationFolderItems: [{}, ...projects.conversationFolderItems] });
  page([hit('channel-a', 'found')], 1);
  expect(await search()).toMatchObject({ data: { skippedItems: 1, results: [{ id: 'found' }] } });
});

it('short-circuits empty sections and distinguishes unavailable membership', async () => {
  hierarchy(folder('p', 'Projects', []));
  expect(await search()).toMatchObject({ data: { results: [], pagination: { hasMore: false, total: 0 } } });
  hierarchy({ id: 'p', name: 'Projects', folderType: 'UserCreated' });
  expect(await search()).toMatchObject({ success: false });
  expect(http).toHaveBeenCalledTimes(2);
});

it.each(['RecentChats', 'TeamsAndChannels'])('reports %s membership as unavailable even with an empty array', async type => {
  const system = folder('system', type, [], type);
  hierarchy(system);
  expect(await invokeTool('teams_get_sections', {})).toMatchObject({ data: { sections: [{ type, itemCount: null }] } });
  hierarchy(system);
  expect(await search({ section: 'system' })).toMatchObject({ success: false, error: { code: ErrorCode.INVALID_INPUT } });
  expect(http).toHaveBeenCalledTimes(2);
});

it('inspects readable members, reports skipped items and moves a chat out of a mixed section', async () => {
  const mixed = { ...projects, conversationFolderItems: [{}, null, { conversationId: 123 }, { conversationId: 'channel-a' }] };
  hierarchy(mixed);
  expect(await invokeTool('teams_get_sections', { section: 'Projects' })).toMatchObject({
    data: { items: [{ conversationId: 'channel-a' }], section: { itemCount: 4 }, skippedItems: 3 },
  });
  hierarchy(mixed, favorites);
  http.mockResolvedValueOnce(response({}));
  expect(await update({ action: 'add', section: 'fav', conversationId: 'channel-a' })).toMatchObject({ success: true });
  expect(body(2).actions).toEqual([
    { action: 'AddItem', folderId: 'fav', itemId: 'channel-a' },
    { action: 'RemoveItem', folderId: 'p/1', itemId: 'channel-a' },
  ]);
});

it('uses the same versioned move for the legacy favorite tools and selects Favorites by type', async () => {
  const sameName = { ...projects, name: 'Favorites' };
  hierarchy(sameName, favorites);
  http.mockResolvedValueOnce(response({}));
  expect(await invokeTool('teams_add_favorite', { conversationId: 'channel-a' })).toMatchObject({
    success: true, data: { message: 'Added channel-a to favourites' },
  });
  expect(body(1)).toMatchObject({ folderHierarchyVersion: 42, actions: [
    { action: 'AddItem', folderId: 'fav', itemId: 'channel-a' },
    { action: 'RemoveItem', folderId: 'p/1', itemId: 'channel-a' },
  ] });
  hierarchy(sameName, favorites);
  http.mockResolvedValueOnce(response({}));
  expect(await invokeTool('teams_remove_favorite', { conversationId: 'chat-b' })).toMatchObject({
    success: true, data: { message: 'Removed chat-b from favourites' },
  });
  expect(body(3).actions).toEqual([{ action: 'RemoveItem', folderId: 'fav', itemId: 'chat-b' }]);
});

it('resumes ordinary search at the first unexamined hit after maxResults', async () => {
  const hits = Array.from({ length: 25 }, (_, i) => hit('channel-a', String(i)));
  page(hits, 25);
  expect(await invokeTool('teams_search', { query: 'launch', maxResults: 5 })).toMatchObject({
    data: { resultCount: 5, pagination: { nextFrom: 5, hasMore: true, total: 25 } },
  });
  page(hits.slice(5), 25);
  expect(await invokeTool('teams_search', { query: 'launch', from: 5 })).toMatchObject({
    data: { resultCount: 20, pagination: { hasMore: false } },
  });
});

it('advances ordinary search past a page of discarded summaries', async () => {
  page([{ Summary: '' }, { Summary: '' }], 3);
  expect(await invokeTool('teams_search', { query: 'launch', size: 2 })).toMatchObject({
    data: { results: [], pagination: { nextFrom: 2, hasMore: true } },
  });
  page([hit('channel-a', 'found')], 3);
  expect(await invokeTool('teams_search', { query: 'launch', from: 2 })).toMatchObject({
    data: { results: [{ id: 'found' }], pagination: { hasMore: false } },
  });
});
