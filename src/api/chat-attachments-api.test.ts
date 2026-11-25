/**
 * Unit tests for chat-message file attachments: local path validation,
 * OneDrive upload + sharing via Microsoft Graph, the chatsvc files-property
 * entry, member resolution, and orchestration.
 *
 * httpRequest and the auth guards are mocked at the network boundary; the
 * filesystem checks run against real temp files under the worktree's .tmp/.
 */

import { describe, it, expect, vi, beforeAll, afterEach } from 'vitest';
import { mkdirSync, writeFileSync, realpathSync } from 'node:fs';
import * as path from 'node:path';
import { httpRequest } from '../utils/http.js';
import { requireGraphTokenAsync, requireMessageAuthWithConfig } from '../utils/auth-guards.js';
import { invalidateAccessToken } from '../auth/token-extractor.js';
import { ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import { chatsvcFilesPropertyEntry } from '../__fixtures__/api-responses.js';
import { MAX_SIMPLE_UPLOAD_BYTES, SELF_CHAT_ID } from '../constants.js';
import {
  resolveUploadFile,
  uploadChatFile,
  shareChatFile,
  buildFilesPropertyEntry,
  getChatMemberObjectIds,
  prepareChatAttachments,
  type UploadedChatFile,
} from './chat-attachments-api.js';

const h = vi.hoisted(() => {
  const root = process.cwd().replace(/\\/g, '/') + '/.tmp/chat-attach-tests';
  return {
    root,
    uploadRoot: root + '/uploads',
    configDir: root + '/uploads/configdir',
    fakeHome: root + '/home',
  };
});

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({
  requireGraphTokenAsync: vi.fn(),
  requireMessageAuthWithConfig: vi.fn(),
}));
vi.mock('../auth/token-extractor.js', () => ({ invalidateAccessToken: vi.fn() }));
vi.mock('../auth/session-store.js', () => ({ CONFIG_DIR: h.configDir }));
vi.mock('node:os', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:os')>();
  return { ...actual, homedir: () => h.fakeHome };
});

const mockHttp = vi.mocked(httpRequest);
const mockGraphToken = vi.mocked(requireGraphTokenAsync);
const mockMessageAuth = vi.mocked(requireMessageAuthWithConfig);

function writeTestFile(rel: string, content = 'hello'): string {
  const file = path.join(h.uploadRoot, rel);
  mkdirSync(path.dirname(file), { recursive: true });
  writeFileSync(file, content);
  return file;
}

beforeAll(() => {
  mkdirSync(h.uploadRoot, { recursive: true });
  mkdirSync(h.fakeHome, { recursive: true });
  mkdirSync(h.configDir, { recursive: true });
});

afterEach(() => {
  vi.unstubAllEnvs();
  vi.clearAllMocks();
});

describe('resolveUploadFile', () => {
  it('accepts a regular file inside the upload root', () => {
    vi.stubEnv('TEAMS_MCP_UPLOAD_DIR', h.uploadRoot);
    const file = writeTestFile('notes.md', 'hello');

    const result = resolveUploadFile(file);

    expect(result).toEqual(ok({
      path: realpathSync(file),
      name: 'notes.md',
      size: 5,
    }));
  });

  it('rejects a relative path', () => {
    vi.stubEnv('TEAMS_MCP_UPLOAD_DIR', h.uploadRoot);
    const file = writeTestFile('rel.md');

    expect(resolveUploadFile(path.relative(process.cwd(), file))).toMatchObject({
      ok: false,
      error: { code: ErrorCode.INVALID_INPUT },
    });
    expect(resolveUploadFile('notes.md')).toMatchObject({
      ok: false,
      error: { code: ErrorCode.INVALID_INPUT },
    });
  });

  it('rejects a file outside the upload root', () => {
    vi.stubEnv('TEAMS_MCP_UPLOAD_DIR', h.uploadRoot);
    const outside = path.join(h.root, 'outside.md');
    writeFileSync(outside, 'secret');

    expect(resolveUploadFile(outside)).toMatchObject({
      ok: false,
      error: { code: ErrorCode.INVALID_INPUT },
    });
    // The upload root itself is not a file inside it.
    expect(resolveUploadFile(h.uploadRoot)).toMatchObject({
      ok: false,
      error: { code: ErrorCode.INVALID_INPUT },
    });
  });

  it('rejects a relative TEAMS_MCP_UPLOAD_DIR', () => {
    vi.stubEnv('TEAMS_MCP_UPLOAD_DIR', 'relative/uploads');
    const file = writeTestFile('env.md');

    expect(resolveUploadFile(file)).toMatchObject({
      ok: false,
      error: { code: ErrorCode.INVALID_INPUT },
    });
  });

  it('rejects dotfiles and dot-directories inside the root', () => {
    vi.stubEnv('TEAMS_MCP_UPLOAD_DIR', h.uploadRoot);
    const dotfile = writeTestFile('.env');
    const inDotDir = writeTestFile('sub/.ssh/id_rsa');
    const deepDot = writeTestFile('sub/.config/app/settings.json');

    for (const file of [dotfile, inDotDir, deepDot]) {
      expect(resolveUploadFile(file)).toMatchObject({
        ok: false,
        error: { code: ErrorCode.INVALID_INPUT },
      });
    }
    // A non-dot sibling is still fine.
    expect(resolveUploadFile(writeTestFile('sub/visible.txt'))).toMatchObject({ ok: true });
  });

  it('rejects anything inside the server config dir', () => {
    vi.stubEnv('TEAMS_MCP_UPLOAD_DIR', h.uploadRoot);
    // The mocked CONFIG_DIR lives inside the upload root so only the
    // config-dir rule can reject this file.
    const secret = path.join(h.configDir, 'session-state.json');
    writeFileSync(secret, '{}');

    expect(resolveUploadFile(secret)).toMatchObject({
      ok: false,
      error: { code: ErrorCode.INVALID_INPUT },
    });
  });

  it.runIf(process.platform === 'win32')('rejects anything under APPDATA or LOCALAPPDATA', () => {
    vi.stubEnv('TEAMS_MCP_UPLOAD_DIR', h.uploadRoot);
    const appData = path.join(h.uploadRoot, 'AppData', 'Roaming');
    const localAppData = path.join(h.uploadRoot, 'AppData', 'Local');
    vi.stubEnv('APPDATA', appData);
    vi.stubEnv('LOCALAPPDATA', localAppData);
    const roaming = path.join(appData, 'SomeApp', 'data.txt');
    const local = path.join(localAppData, 'SomeApp', 'cache.bin');
    mkdirSync(path.dirname(roaming), { recursive: true });
    mkdirSync(path.dirname(local), { recursive: true });
    writeFileSync(roaming, 'x');
    writeFileSync(local, 'x');

    for (const file of [roaming, local]) {
      expect(resolveUploadFile(file)).toMatchObject({
        ok: false,
        error: { code: ErrorCode.INVALID_INPUT },
      });
    }
  });

  it('rejects directories and missing files', () => {
    vi.stubEnv('TEAMS_MCP_UPLOAD_DIR', h.uploadRoot);
    mkdirSync(path.join(h.uploadRoot, 'adir'), { recursive: true });

    expect(resolveUploadFile(path.join(h.uploadRoot, 'adir'))).toMatchObject({
      ok: false,
      error: { code: ErrorCode.INVALID_INPUT },
    });
    expect(resolveUploadFile(path.join(h.uploadRoot, 'no-such-file.txt'))).toMatchObject({
      ok: false,
      error: { code: ErrorCode.INVALID_INPUT },
    });
  });

  it('rejects files larger than the simple-upload limit', () => {
    vi.stubEnv('TEAMS_MCP_UPLOAD_DIR', h.uploadRoot);
    const big = path.join(h.uploadRoot, 'big.bin');
    writeFileSync(big, Buffer.alloc(MAX_SIMPLE_UPLOAD_BYTES + 1));
    const atLimit = path.join(h.uploadRoot, 'at-limit.bin');
    writeFileSync(atLimit, Buffer.alloc(MAX_SIMPLE_UPLOAD_BYTES));

    expect(resolveUploadFile(big)).toMatchObject({
      ok: false,
      error: { code: ErrorCode.INVALID_INPUT },
    });
    expect(resolveUploadFile(atLimit)).toMatchObject({ ok: true });
  });

  it('defaults the upload root to the home directory', () => {
    vi.stubEnv('TEAMS_MCP_UPLOAD_DIR', undefined);
    const home = path.join(h.fakeHome, 'desktop', 'note.txt');
    mkdirSync(path.dirname(home), { recursive: true });
    writeFileSync(home, 'hi');
    const elsewhere = writeTestFile('not-home.txt');

    expect(resolveUploadFile(home)).toMatchObject({ ok: true, value: { name: 'note.txt' } });
    expect(resolveUploadFile(elsewhere)).toMatchObject({
      ok: false,
      error: { code: ErrorCode.INVALID_INPUT },
    });
  });
});

describe('uploadChatFile', () => {
  const DRIVE_ITEM = {
    id: '01DRIVEITEMID',
    name: 'notes.md',
    webUrl: 'https://contoso-my.sharepoint.com/personal/me/Documents/Microsoft%20Teams%20Chat%20Files/notes.md',
    sharepointIds: {
      listId: 'list-1',
      listItemUniqueId: 'item-1',
      siteId: 'site-1',
      siteUrl: 'https://contoso-my.sharepoint.com/personal/me',
      webId: 'web-1',
    },
    parentReference: { driveId: 'b!drive1' },
  };

  function localFile(): { path: string; name: string; size: number } {
    vi.stubEnv('TEAMS_MCP_UPLOAD_DIR', h.uploadRoot);
    const file = writeTestFile('up/notes.md', 'file bytes');
    const resolved = resolveUploadFile(file);
    if (!resolved.ok) throw new Error('test setup: resolveUploadFile failed');
    return resolved.value;
  }

  it('PUTs the bytes to the Teams Chat Files folder, then reads the item back', async () => {
    const file = localFile();
    mockHttp
      .mockResolvedValueOnce(ok({ status: 201, headers: new Headers(), data: { id: '01DRIVEITEMID' } }))
      .mockResolvedValueOnce(ok({ status: 200, headers: new Headers(), data: DRIVE_ITEM }));

    const result = await uploadChatFile(file, 'graph-token');

    expect(result).toEqual(ok({
      id: '01DRIVEITEMID',
      name: 'notes.md',
      webUrl: DRIVE_ITEM.webUrl,
      driveId: 'b!drive1',
      sharepointIds: DRIVE_ITEM.sharepointIds,
    }));

    const [putUrl, putOpts] = mockHttp.mock.calls[0];
    const put = new URL(putUrl as string);
    expect(put.origin).toBe('https://graph.microsoft.com');
    expect(put.pathname).toBe('/v1.0/me/drive/root:/Microsoft%20Teams%20Chat%20Files/notes.md:/content');
    expect(put.searchParams.get('@microsoft.graph.conflictBehavior')).toBe('rename');
    expect(putOpts).toMatchObject({
      method: 'PUT',
      maxRetries: 1,
      headers: expect.objectContaining({
        Authorization: 'Bearer graph-token',
        'Content-Type': 'application/octet-stream',
      }),
    });
    expect((putOpts as { body: Buffer }).body).toEqual(Buffer.from('file bytes'));

    const [getUrl, getOpts] = mockHttp.mock.calls[1];
    const get = new URL(getUrl as string);
    expect(get.pathname).toBe('/v1.0/me/drive/items/01DRIVEITEMID');
    expect(get.searchParams.get('$select')).toContain('sharepointIds');
    expect(get.searchParams.get('$select')).toContain('parentReference');
    expect((getOpts as { headers: Record<string, string> }).headers.Authorization).toBe('Bearer graph-token');
  });

  it('stops when the PUT fails and never reads the item back', async () => {
    const file = localFile();
    mockHttp.mockResolvedValueOnce(err(createError(ErrorCode.API_ERROR, 'HTTP 500: boom')));

    const result = await uploadChatFile(file, 'graph-token');

    expect(result).toMatchObject({ ok: false, error: { code: ErrorCode.API_ERROR } });
    expect(mockHttp).toHaveBeenCalledTimes(1);
  });

  it('fails when the PUT response has no drive item id', async () => {
    const file = localFile();
    mockHttp.mockResolvedValueOnce(ok({ status: 200, headers: new Headers(), data: {} }));

    expect(await uploadChatFile(file, 'graph-token')).toMatchObject({
      ok: false,
      error: { code: ErrorCode.API_ERROR },
    });
    expect(mockHttp).toHaveBeenCalledTimes(1);
  });

  it('fails when the drive item lacks fields chatsvc needs', async () => {
    const file = localFile();
    mockHttp
      .mockResolvedValueOnce(ok({ status: 201, headers: new Headers(), data: { id: '01DRIVEITEMID' } }))
      .mockResolvedValueOnce(ok({ status: 200, headers: new Headers(), data: { ...DRIVE_ITEM, webUrl: undefined } }));

    expect(await uploadChatFile(file, 'graph-token')).toMatchObject({
      ok: false,
      error: { code: ErrorCode.API_ERROR },
    });
  });

  it('remaps a rejected Graph token without triggering Teams re-login', async () => {
    const file = localFile();
    mockHttp.mockResolvedValueOnce(err(createError(ErrorCode.AUTH_EXPIRED, 'HTTP 401')));

    expect(await uploadChatFile(file, 'graph-token')).toMatchObject({
      ok: false,
      error: { code: ErrorCode.API_ERROR, retryable: true },
    });
    expect(invalidateAccessToken).toHaveBeenCalledWith('graph-token');
  });

  it('reports a Graph 403 as access denied, not an auth error', async () => {
    const file = localFile();
    mockHttp.mockResolvedValueOnce(err(createError(ErrorCode.AUTH_REQUIRED, 'HTTP 403')));

    expect(await uploadChatFile(file, 'graph-token')).toMatchObject({
      ok: false,
      error: { code: ErrorCode.ACCESS_DENIED },
    });
  });
});

describe('shareChatFile', () => {
  const WEB_URL = 'https://contoso.sharepoint.com/:t:/g/personal/me/FAKESHARETOKEN';

  it('creates a view link for users, then grants read to the recipients', async () => {
    mockHttp
      .mockResolvedValueOnce(ok({ status: 200, headers: new Headers(), data: { id: 'perm-1', link: { webUrl: WEB_URL } } }))
      .mockResolvedValueOnce(ok({ status: 200, headers: new Headers(), data: {} }));

    const result = await shareChatFile('01ITEM', ['oid-1', 'oid-2'], 'graph-token');

    expect(result).toEqual(ok({ shareId: 'perm-1', shareUrl: WEB_URL }));

    const [linkUrl, linkOpts] = mockHttp.mock.calls[0];
    expect(new URL(linkUrl as string).pathname).toBe('/v1.0/me/drive/items/01ITEM/createLink');
    expect(linkOpts).toMatchObject({
      method: 'POST',
      maxRetries: 1,
      body: JSON.stringify({ type: 'view', scope: 'users' }),
      headers: expect.objectContaining({ Authorization: 'Bearer graph-token' }),
    });

    const [grantUrl, grantOpts] = mockHttp.mock.calls[1];
    const grant = new URL(grantUrl as string);
    expect(grant.origin).toBe('https://graph.microsoft.com');
    expect(grant.pathname).toBe(`/v1.0/shares/u!${Buffer.from(WEB_URL).toString('base64url')}/permission/grant`);
    expect(grantOpts).toMatchObject({
      method: 'POST',
      maxRetries: 1,
      body: JSON.stringify({
        recipients: [{ objectId: 'oid-1' }, { objectId: 'oid-2' }],
        roles: ['read'],
      }),
    });
  });

  it('skips the grant call when there are no recipients', async () => {
    mockHttp.mockResolvedValueOnce(ok({
      status: 200, headers: new Headers(), data: { id: 'perm-1', link: { webUrl: WEB_URL } },
    }));

    const result = await shareChatFile('01ITEM', [], 'graph-token');

    expect(result).toEqual(ok({ shareId: 'perm-1', shareUrl: WEB_URL }));
    expect(mockHttp).toHaveBeenCalledTimes(1);
    expect(mockHttp.mock.calls[0][0]).toContain('/createLink');
  });

  it('fails without granting when createLink fails', async () => {
    mockHttp.mockResolvedValueOnce(err(createError(ErrorCode.API_ERROR, 'HTTP 500')));

    expect(await shareChatFile('01ITEM', ['oid-1'], 'graph-token')).toMatchObject({
      ok: false,
      error: { code: ErrorCode.API_ERROR },
    });
    expect(mockHttp).toHaveBeenCalledTimes(1);
  });

  it('fails when createLink returns no sharing link', async () => {
    mockHttp.mockResolvedValueOnce(ok({ status: 200, headers: new Headers(), data: { id: 'perm-1' } }));

    expect(await shareChatFile('01ITEM', ['oid-1'], 'graph-token')).toMatchObject({
      ok: false,
      error: { code: ErrorCode.API_ERROR },
    });
    expect(mockHttp).toHaveBeenCalledTimes(1);
  });

  it('fails when the grant call fails', async () => {
    mockHttp
      .mockResolvedValueOnce(ok({ status: 200, headers: new Headers(), data: { id: 'perm-1', link: { webUrl: WEB_URL } } }))
      .mockResolvedValueOnce(err(createError(ErrorCode.API_ERROR, 'HTTP 500')));

    expect(await shareChatFile('01ITEM', ['oid-1'], 'graph-token')).toMatchObject({
      ok: false,
      error: { code: ErrorCode.API_ERROR },
    });
    expect(mockHttp).toHaveBeenCalledTimes(2);
  });

  it('remaps a rejected Graph token on the grant call', async () => {
    mockHttp
      .mockResolvedValueOnce(ok({ status: 200, headers: new Headers(), data: { id: 'perm-1', link: { webUrl: WEB_URL } } }))
      .mockResolvedValueOnce(err(createError(ErrorCode.AUTH_EXPIRED, 'HTTP 401')));

    expect(await shareChatFile('01ITEM', ['oid-1'], 'graph-token')).toMatchObject({
      ok: false,
      error: { code: ErrorCode.API_ERROR },
    });
    expect(invalidateAccessToken).toHaveBeenCalledWith('graph-token');
  });
});

describe('buildFilesPropertyEntry', () => {
  // Values chosen so every mapped field lands exactly on the fixture's.
  const uploaded: UploadedChatFile = {
    id: '01FAKEDRIVEITEMID0000000000000000AB',
    name: 'design-notes.md',
    webUrl: 'https://contoso-my.sharepoint.com/personal/jane_doe_contoso_com/Documents/Microsoft%20Teams%20Chat%20Files/design-notes.md',
    driveId: 'b!FAKEDRIVEID000000000000000000000000000000000000000000000000000',
    sharepointIds: {
      listId: 'bbbbbbbb-cccc-4ddd-8eee-ffffffffffff',
      listItemUniqueId: '11111111-2222-4333-8444-555555555555',
      siteId: '12121212-3434-4565-8787-909090909090',
      siteUrl: 'https://contoso-my.sharepoint.com/personal/jane_doe_contoso_com',
      webId: '13131313-2424-4535-8646-757575757575',
    },
  };
  const share = {
    shareId: '66666666-7777-4888-9999-aaaaaaaaaaaa',
    shareUrl: 'https://contoso-my.sharepoint.com/:t:/g/personal/jane_doe_contoso_com/FAKESHARETOKEN0000000000000000000000000000000',
  };

  it('produces an entry identical to a real chatsvc files property, key order included', () => {
    const entry = buildFilesPropertyEntry(uploaded, share);

    expect(entry).toEqual(chatsvcFilesPropertyEntry);
    // properties.files is JSON-encoded on the wire, so key order is observable.
    expect(Object.keys(entry)).toEqual(Object.keys(chatsvcFilesPropertyEntry));
    expect(Object.keys(entry.fileInfo)).toEqual(Object.keys(chatsvcFilesPropertyEntry.fileInfo));
    expect(Object.keys(entry.fileChicletState)).toEqual(Object.keys(chatsvcFilesPropertyEntry.fileChicletState));
    expect(Object.keys(entry.filePreview)).toEqual(Object.keys(chatsvcFilesPropertyEntry.filePreview));
    expect(Object.keys(entry.sharepointIds)).toEqual(Object.keys(chatsvcFilesPropertyEntry.sharepointIds));
  });

  it('derives fileType/type from the Graph name extension', () => {
    const noExt = buildFilesPropertyEntry({ ...uploaded, name: 'README' }, share);
    expect(noExt.fileType).toBe('');
    expect(noExt.type).toBe('');
    expect(noExt.fileName).toBe('README');
    expect(noExt.title).toBe('README');

    const upper = buildFilesPropertyEntry({ ...uploaded, name: 'NOTES.MD' }, share);
    expect(upper.fileType).toBe('md');
    expect(upper.type).toBe('md');

    const tarball = buildFilesPropertyEntry({ ...uploaded, name: 'archive.tar.gz' }, share);
    expect(tarball.fileType).toBe('gz');
  });
});

describe('getChatMemberObjectIds', () => {
  const CONVERSATION = '19:abc@thread.v2';
  const SELF_MRI = '8:orgid:99999999-8888-7777-6666-555555555555';

  function mockChatAuth() {
    mockMessageAuth.mockReturnValue(ok({
      auth: { skypeToken: 'sk', authToken: 'at', userMri: SELF_MRI },
      region: 'amer',
      baseUrl: 'https://teams.microsoft.com',
    }));
  }

  it('returns the other members object IDs, excluding self and non-users', async () => {
    mockChatAuth();
    mockHttp.mockResolvedValueOnce(ok({
      status: 200,
      headers: new Headers(),
      data: {
        members: [
          { id: '8:orgid:11111111-2222-3333-4444-555555555555' },
          { id: SELF_MRI },
          { id: '8:teams:bot-id' },
          { id: '8:orgid:11111111-2222-3333-4444-555555555555' },
          { id: '8:orgid:AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE' },
          { id: '8:orgid:not-a-guid' },
          { displayName: 'no id' },
        ],
      },
    }));

    const result = await getChatMemberObjectIds(CONVERSATION);

    expect(result).toEqual(ok([
      '11111111-2222-3333-4444-555555555555',
      'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
    ]));

    const [url, opts] = mockHttp.mock.calls[0];
    const u = new URL(url as string);
    expect(u.pathname).toBe('/api/chatsvc/amer/v1/threads/19%3Aabc%40thread.v2');
    expect(u.searchParams.get('view')).toBe('msnp24Equivalent');
    expect(opts).toMatchObject({
      method: 'GET',
      headers: expect.objectContaining({
        Authentication: 'skypetoken=sk',
        Authorization: 'Bearer at',
      }),
    });
  });

  it('passes core-auth failures through unchanged', async () => {
    const denied = err(createError(ErrorCode.AUTH_REQUIRED, 'no session'));
    mockMessageAuth.mockReturnValue(denied);
    expect(await getChatMemberObjectIds(CONVERSATION)).toEqual(denied);
    expect(mockHttp).not.toHaveBeenCalled();

    mockChatAuth();
    mockHttp.mockResolvedValueOnce(err(createError(ErrorCode.AUTH_EXPIRED, 'HTTP 401')));
    expect(await getChatMemberObjectIds(CONVERSATION)).toMatchObject({
      ok: false,
      error: { code: ErrorCode.AUTH_EXPIRED },
    });
  });

  it('fails when the thread response has no members array', async () => {
    mockChatAuth();
    mockHttp.mockResolvedValueOnce(ok({ status: 200, headers: new Headers(), data: {} }));

    expect(await getChatMemberObjectIds(CONVERSATION)).toMatchObject({
      ok: false,
      error: { code: ErrorCode.API_ERROR },
    });
  });
});

describe('prepareChatAttachments', () => {
  const CHAT = '19:abc@thread.v2';

  // Values mirror the buildFilesPropertyEntry fixtures so a fully mocked run
  // produces exactly chatsvcFilesPropertyEntry.
  const DRIVE_ITEM = {
    id: '01FAKEDRIVEITEMID0000000000000000AB',
    name: 'design-notes.md',
    webUrl: 'https://contoso-my.sharepoint.com/personal/jane_doe_contoso_com/Documents/Microsoft%20Teams%20Chat%20Files/design-notes.md',
    sharepointIds: {
      listId: 'bbbbbbbb-cccc-4ddd-8eee-ffffffffffff',
      listItemUniqueId: '11111111-2222-4333-8444-555555555555',
      siteId: '12121212-3434-4565-8787-909090909090',
      siteUrl: 'https://contoso-my.sharepoint.com/personal/jane_doe_contoso_com',
      webId: '13131313-2424-4535-8646-757575757575',
    },
    parentReference: { driveId: 'b!FAKEDRIVEID000000000000000000000000000000000000000000000000000' },
  };
  const SHARE_PERMISSION = {
    id: '66666666-7777-4888-9999-aaaaaaaaaaaa',
    link: { webUrl: 'https://contoso-my.sharepoint.com/:t:/g/personal/jane_doe_contoso_com/FAKESHARETOKEN0000000000000000000000000000000' },
  };

  function mockGraphOk() {
    mockGraphToken.mockResolvedValue(ok('graph-token'));
  }

  function mockChatAuth() {
    mockMessageAuth.mockReturnValue(ok({
      auth: { skypeToken: 'sk', authToken: 'at', userMri: '8:orgid:99999999-8888-7777-6666-555555555555' },
      region: 'amer',
      baseUrl: 'https://teams.microsoft.com',
    }));
  }

  it('validates every path before any network call', async () => {
    vi.stubEnv('TEAMS_MCP_UPLOAD_DIR', h.uploadRoot);
    const good = writeTestFile('prep/good.txt');
    const bad = path.join(h.uploadRoot, 'prep', 'missing.txt');

    const result = await prepareChatAttachments(CHAT, [good, bad]);

    expect(result).toMatchObject({ ok: false, error: { code: ErrorCode.INVALID_INPUT } });
    expect(mockHttp).not.toHaveBeenCalled();
    expect(mockGraphToken).not.toHaveBeenCalled();
    expect(mockMessageAuth).not.toHaveBeenCalled();

    expect(await prepareChatAttachments(CHAT, [])).toEqual(ok([]));
  });

  it('rejects channel conversations before any network call', async () => {
    vi.stubEnv('TEAMS_MCP_UPLOAD_DIR', h.uploadRoot);
    const file = writeTestFile('prep/note.txt');

    for (const id of ['19:chan@thread.tacv2', '19:chan@thread.skype']) {
      const result = await prepareChatAttachments(id, [file]);
      expect(result).toMatchObject({ ok: false, error: { code: ErrorCode.INVALID_INPUT } });
      if (!result.ok) {
        expect(result.error.message).toContain('channel');
      }
    }
    expect(mockHttp).not.toHaveBeenCalled();
    expect(mockGraphToken).not.toHaveBeenCalled();
    expect(mockMessageAuth).not.toHaveBeenCalled();
  });

  it('fails before any upload when Graph auth is unavailable', async () => {
    vi.stubEnv('TEAMS_MCP_UPLOAD_DIR', h.uploadRoot);
    const file = writeTestFile('prep/note2.txt');
    mockGraphToken.mockResolvedValue(err(createError(ErrorCode.ACCESS_DENIED, 'no consent')));

    const result = await prepareChatAttachments(CHAT, [file]);

    expect(result).toMatchObject({ ok: false, error: { code: ErrorCode.ACCESS_DENIED } });
    expect(mockHttp).not.toHaveBeenCalled();
    expect(mockMessageAuth).not.toHaveBeenCalled();
  });

  it('uploads and shares each file for self-notes without a members lookup', async () => {
    vi.stubEnv('TEAMS_MCP_UPLOAD_DIR', h.uploadRoot);
    const file = writeTestFile('prep/design-notes.md', 'notes');
    mockGraphOk();
    mockHttp
      .mockResolvedValueOnce(ok({ status: 201, headers: new Headers(), data: { id: DRIVE_ITEM.id } }))
      .mockResolvedValueOnce(ok({ status: 200, headers: new Headers(), data: DRIVE_ITEM }))
      .mockResolvedValueOnce(ok({ status: 200, headers: new Headers(), data: SHARE_PERMISSION }));

    const result = await prepareChatAttachments(SELF_CHAT_ID, [file]);

    expect(result).toEqual(ok([chatsvcFilesPropertyEntry]));
    // Self-notes have no other members: no chatsvc call, no grant call.
    expect(mockMessageAuth).not.toHaveBeenCalled();
    expect(mockHttp).toHaveBeenCalledTimes(3);
  });

  it('shares with the other members of a group chat', async () => {
    vi.stubEnv('TEAMS_MCP_UPLOAD_DIR', h.uploadRoot);
    const file = writeTestFile('prep/group.md');
    mockGraphOk();
    mockChatAuth();
    mockHttp
      .mockResolvedValueOnce(ok({
        status: 200,
        headers: new Headers(),
        data: { members: [{ id: '8:orgid:11111111-2222-3333-4444-555555555555' }] },
      }))
      .mockResolvedValueOnce(ok({ status: 201, headers: new Headers(), data: { id: DRIVE_ITEM.id } }))
      .mockResolvedValueOnce(ok({ status: 200, headers: new Headers(), data: DRIVE_ITEM }))
      .mockResolvedValueOnce(ok({ status: 200, headers: new Headers(), data: SHARE_PERMISSION }))
      .mockResolvedValueOnce(ok({ status: 200, headers: new Headers(), data: {} }));

    const result = await prepareChatAttachments(CHAT, [file]);

    expect(result.ok).toBe(true);
    // threads lookup + PUT + GET + createLink + grant
    expect(mockHttp).toHaveBeenCalledTimes(5);
    const grantBody = JSON.parse((mockHttp.mock.calls[4][1] as { body: string }).body);
    expect(grantBody).toEqual({
      recipients: [{ objectId: '11111111-2222-3333-4444-555555555555' }],
      roles: ['read'],
    });
    expect(mockHttp.mock.calls[4][0]).toContain('/permission/grant');
  });

  it('names the failing file and reports that nothing was sent', async () => {
    vi.stubEnv('TEAMS_MCP_UPLOAD_DIR', h.uploadRoot);
    const first = writeTestFile('prep/first.txt', '1');
    const second = writeTestFile('prep/second.txt', '2');
    mockGraphOk();
    mockHttp
      .mockResolvedValueOnce(ok({ status: 201, headers: new Headers(), data: { id: DRIVE_ITEM.id } }))
      .mockResolvedValueOnce(ok({ status: 200, headers: new Headers(), data: DRIVE_ITEM }))
      .mockResolvedValueOnce(ok({ status: 200, headers: new Headers(), data: SHARE_PERMISSION }))
      .mockResolvedValueOnce(err(createError(ErrorCode.API_ERROR, 'HTTP 500: boom')));

    const result = await prepareChatAttachments(SELF_CHAT_ID, [first, second]);

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe(ErrorCode.API_ERROR);
      expect(result.error.message).toContain('second.txt');
      expect(result.error.message).toMatch(/nothing was sent/i);
    }
    // First file completed its 3 calls; the second file's PUT failed and
    // processing stopped there.
    expect(mockHttp).toHaveBeenCalledTimes(4);
    expect(mockHttp.mock.calls[3][0]).toContain('second.txt');
  });

  it('passes member-lookup auth failures through for group chats', async () => {
    vi.stubEnv('TEAMS_MCP_UPLOAD_DIR', h.uploadRoot);
    const file = writeTestFile('prep/x.txt');
    mockGraphOk();
    mockChatAuth();
    mockHttp.mockResolvedValueOnce(err(createError(ErrorCode.AUTH_EXPIRED, 'HTTP 401')));

    const result = await prepareChatAttachments(CHAT, [file]);

    expect(result).toMatchObject({ ok: false, error: { code: ErrorCode.AUTH_EXPIRED } });
    // Only the chatsvc threads call ran; no Graph upload was attempted.
    expect(mockHttp).toHaveBeenCalledTimes(1);
    expect(mockHttp.mock.calls[0][0]).toContain('/api/chatsvc/');
  });
});
