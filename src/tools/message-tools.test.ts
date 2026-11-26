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
vi.mock('../api/chat-attachments-api.js', () => ({
  prepareChatAttachments: vi.fn(),
  clearPreparedAttachments: vi.fn(),
}));

import { sendMessageTool } from './message-tools.js';
import { sendMessage } from '../api/chatsvc-api.js';
import { prepareChatAttachments, clearPreparedAttachments } from '../api/chat-attachments-api.js';

const mockSendMessage = vi.mocked(sendMessage);
const mockPrepare = vi.mocked(prepareChatAttachments);
const mockClear = vi.mocked(clearPreparedAttachments);

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
    expect(mockClear).not.toHaveBeenCalled();
  });

  it('clears the prepared attachments after a successful send', async () => {
    const result = await sendMessageTool.handler({
      content: 'with file',
      conversationId: '19:abc@thread.v2',
      files: ['/tmp/a.txt'],
    });

    expect(result.success).toBe(true);
    expect(mockClear).toHaveBeenCalledWith('19:abc@thread.v2', ['/tmp/a.txt']);
  });

  it('does not clear prepared attachments when the send fails, and says the files are reusable', async () => {
    mockSendMessage.mockResolvedValue(err(createError(ErrorCode.API_ERROR, 'chatsvc HTTP 500', {
      retryable: true,
      suggestions: ['Try again later'],
    })));

    const result = await sendMessageTool.handler({
      content: 'with file',
      conversationId: '19:abc@thread.v2',
      files: ['/tmp/a.txt'],
    });

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe(ErrorCode.API_ERROR);
      expect(result.error.retryable).toBe(true);
      expect(result.error.suggestions).toEqual(['Try again later']);
      expect(result.error.message).toContain('chatsvc HTTP 500');
      expect(result.error.message).toContain(
        '1 file(s) were uploaded to your OneDrive "Microsoft Teams Chat Files" folder and shared, but the message was not sent; retrying within 10 minutes reuses them.');
    }
    // The cache must survive so the server's auto-retry reuses the uploads.
    expect(mockClear).not.toHaveBeenCalled();
  });

  it.each([
    ['a timeout', ErrorCode.TIMEOUT, 'Request timed out after 30000ms'],
    ['a network error', ErrorCode.NETWORK_ERROR, 'read ECONNRESET'],
    ['a 5xx response', ErrorCode.API_ERROR, 'HTTP 503: upstream unavailable'],
  ])('says the message may have been delivered on %s', async (_label, code, message) => {
    mockSendMessage.mockResolvedValue(err(createError(code, message, { retryable: true })));

    const result = await sendMessageTool.handler({
      content: 'with file',
      conversationId: '19:abc@thread.v2',
      files: ['/tmp/a.txt'],
    });

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.message).toContain('may or may not have been delivered');
      expect(result.error.message).toContain('teams_get_thread');
      expect(result.error.message).not.toContain('was not sent');
      expect(result.error.message).toContain('10 minutes');
    }
    expect(mockClear).not.toHaveBeenCalled();
  });

  it('keeps "was not sent" for a definite send rejection', async () => {
    mockSendMessage.mockResolvedValue(err(createError(ErrorCode.INVALID_INPUT, 'bad conversation')));

    const result = await sendMessageTool.handler({
      content: 'with file',
      conversationId: '19:abc@thread.v2',
      files: ['/tmp/a.txt'],
    });

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.message).toContain('was not sent');
      expect(result.error.message).not.toContain('may or may not');
    }
    expect(mockClear).not.toHaveBeenCalled();
  });
});
