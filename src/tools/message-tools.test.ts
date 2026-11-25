/**
 * Unit tests for the teams_send_message tool handler's files option:
 * attachments are prepared via chat-attachments-api, then passed to
 * sendMessage as chatsvc properties.files entries.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import { chatsvcFilesPropertyEntry } from '../__fixtures__/api-responses.js';

vi.mock('../api/chatsvc-api.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../api/chatsvc-api.js')>();
  return { ...actual, sendMessage: vi.fn() };
});
vi.mock('../api/chat-attachments-api.js', () => ({ prepareChatAttachments: vi.fn() }));

import { sendMessageTool } from './message-tools.js';
import { sendMessage } from '../api/chatsvc-api.js';
import { prepareChatAttachments } from '../api/chat-attachments-api.js';

const mockSendMessage = vi.mocked(sendMessage);
const mockPrepare = vi.mocked(prepareChatAttachments);

const ENTRIES = [chatsvcFilesPropertyEntry];

beforeEach(() => {
  vi.clearAllMocks();
  mockSendMessage.mockResolvedValue(ok({ messageId: 'cmid', timestamp: 123 }));
  mockPrepare.mockResolvedValue(ok(ENTRIES));
});

describe('teams_send_message files', () => {
  it('rejects files combined with scheduleAt before preparing or sending', async () => {
    const result = await sendMessageTool.handler({
      content: 'later',
      conversationId: '19:abc@thread.v2',
      files: ['/tmp/a.txt'],
      scheduleAt: '2099-12-31T09:00:00Z',
    });

    expect(result).toMatchObject({ success: false, error: { code: ErrorCode.INVALID_INPUT } });
    expect(mockPrepare).not.toHaveBeenCalled();
    expect(mockSendMessage).not.toHaveBeenCalled();
  });

  it('returns the prepare error and never sends when attachment prep fails', async () => {
    mockPrepare.mockResolvedValue(err(createError(ErrorCode.API_ERROR, 'Graph is down')));

    const result = await sendMessageTool.handler({
      content: 'with file',
      conversationId: '19:abc@thread.v2',
      files: ['/tmp/a.txt'],
    });

    expect(result).toMatchObject({ success: false, error: { code: ErrorCode.API_ERROR } });
    expect(mockPrepare).toHaveBeenCalledWith('19:abc@thread.v2', ['/tmp/a.txt']);
    expect(mockSendMessage).not.toHaveBeenCalled();
  });

  it('passes the prepared entries to sendMessage', async () => {
    const result = await sendMessageTool.handler({
      content: 'with file',
      conversationId: '19:abc@thread.v2',
      files: ['/tmp/a.txt', '/tmp/b.txt'],
    });

    expect(result.success).toBe(true);
    expect(mockPrepare).toHaveBeenCalledWith('19:abc@thread.v2', ['/tmp/a.txt', '/tmp/b.txt']);
    expect(mockSendMessage).toHaveBeenCalledWith('19:abc@thread.v2', 'with file',
      expect.objectContaining({ files: ENTRIES }));
  });

  it('does not prepare attachments when files is absent', async () => {
    const result = await sendMessageTool.handler({
      content: 'plain',
      conversationId: '19:abc@thread.v2',
    });

    expect(result.success).toBe(true);
    expect(mockPrepare).not.toHaveBeenCalled();
    expect(mockSendMessage).toHaveBeenCalledWith('19:abc@thread.v2', 'plain',
      expect.objectContaining({ files: undefined }));
  });
});
