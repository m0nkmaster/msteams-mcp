/**
 * Unit tests for sendMessage's HTTP body, focused on the files option that
 * carries chatsvc properties.files entries for chat attachments.
 *
 * httpRequest and the auth guards are mocked at the network boundary.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { httpRequest } from '../utils/http.js';
import { requireMessageAuthWithConfig } from '../utils/auth-guards.js';
import { ok } from '../types/result.js';
import { ErrorCode } from '../types/errors.js';
import { chatsvcFilesPropertyEntry } from '../__fixtures__/api-responses.js';
import { sendMessage } from './chatsvc-messaging.js';
import type { ChatsvcFileEntry } from './chat-attachments-api.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({
  requireMessageAuth: vi.fn(),
  requireMessageAuthWithConfig: vi.fn(),
  getTeamsBaseUrl: vi.fn(),
  getTenantId: vi.fn(),
}));
vi.mock('../auth/token-extractor.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../auth/token-extractor.js')>();
  return { ...actual, getUserDisplayName: vi.fn(() => 'Test User') };
});

const mockHttp = vi.mocked(httpRequest);
const mockMessageAuth = vi.mocked(requireMessageAuthWithConfig);

const CONVERSATION = '19:abc@thread.v2';
const FILES = [chatsvcFilesPropertyEntry] as unknown as ChatsvcFileEntry[];

function mockAuth() {
  mockMessageAuth.mockReturnValue(ok({
    auth: { skypeToken: 'sk', authToken: 'at', userMri: '8:orgid:me' },
    region: 'amer',
    baseUrl: 'https://teams.microsoft.com',
  }));
}

function lastBody(): Record<string, unknown> {
  const [, opts] = mockHttp.mock.calls[mockHttp.mock.calls.length - 1];
  return JSON.parse((opts as { body: string }).body) as Record<string, unknown>;
}

beforeEach(() => {
  vi.clearAllMocks();
  mockAuth();
  mockHttp.mockResolvedValue(ok({ status: 200, headers: new Headers(), data: { OriginalArrivalTime: 123 } }));
});

describe('sendMessage files option', () => {
  it('serialises files entries into properties.files as a JSON string', async () => {
    const result = await sendMessage(CONVERSATION, 'here is the file', { files: FILES });

    expect(result.ok).toBe(true);
    const body = lastBody();
    const properties = body.properties as Record<string, unknown>;
    expect(properties.files).toBe(JSON.stringify(FILES));
    expect(JSON.parse(properties.files as string)).toEqual(FILES);
  });

  it('merges files with mentions and subject in properties', async () => {
    const result = await sendMessage(
      CONVERSATION,
      'Hi @[Alice](8:orgid:abc), file attached',
      { files: FILES, subject: 'Q3 report' },
    );

    expect(result.ok).toBe(true);
    const properties = lastBody().properties as Record<string, unknown>;
    expect(properties.files).toBe(JSON.stringify(FILES));
    expect(properties.subject).toBe('Q3 report');
    expect(JSON.parse(properties.mentions as string)).toHaveLength(1);
  });

  it('omits properties.files for an empty files array', async () => {
    const result = await sendMessage(CONVERSATION, 'plain', { files: [] });

    expect(result.ok).toBe(true);
    expect(lastBody().properties).toBeUndefined();
  });

  it('rejects files combined with scheduleAt before any request', async () => {
    const result = await sendMessage(CONVERSATION, 'later', {
      files: FILES,
      scheduleAt: '2099-12-31T09:00:00Z',
    });

    expect(result).toMatchObject({ ok: false, error: { code: ErrorCode.INVALID_INPUT } });
    expect(mockHttp).not.toHaveBeenCalled();
  });
});
