/**
 * Unit tests for conversation-list parsing.
 */

import { describe, it, expect } from 'vitest';
import { parseConversation } from './chatsvc-conversations.js';

describe('parseConversation', () => {
  it('parses a 1:1 chat with a last-message preview', () => {
    const result = parseConversation({
      id: '19:aaa_bbb@unq.gbl.spaces',
      lastMessage: {
        imdisplayname: 'Bob',
        content: '<p>Hi there</p>',
        originalarrivaltime: '2026-06-30T10:00:00Z',
      },
    });
    expect(result).toEqual({
      conversationId: '19:aaa_bbb@unq.gbl.spaces',
      topic: '',
      chatType: 'oneOnOne',
      lastMessageTime: '2026-06-30T10:00:00Z',
      lastMessageFrom: 'Bob',
      lastMessagePreview: 'Hi there',
    });
  });

  it('classifies channel, group and meeting conversation IDs', () => {
    expect(parseConversation({ id: '19:x@thread.tacv2', threadProperties: { topic: 'General' } })?.chatType).toBe('channel');
    expect(parseConversation({ id: '19:x@thread.v2' })?.chatType).toBe('group');
    expect(parseConversation({ id: '19:meeting_x@thread.v2' })?.chatType).toBe('meeting');
  });

  it('uses the channel/group topic when present', () => {
    const result = parseConversation({ id: '19:x@thread.tacv2', threadProperties: { topic: 'General' } });
    expect(result?.topic).toBe('General');
  });

  it('truncates a long preview to 120 characters with an ellipsis', () => {
    const long = 'a'.repeat(200);
    const result = parseConversation({ id: '19:x@thread.v2', lastMessage: { content: long } });
    expect(result?.lastMessagePreview.length).toBe(121); // 120 chars + ellipsis
    expect(result?.lastMessagePreview.endsWith('…')).toBe(true);
  });

  it('returns null when the conversation has no id', () => {
    expect(parseConversation({})).toBeNull();
    expect(parseConversation({ id: '' })).toBeNull();
  });

  it('falls back to composetime when originalarrivaltime is absent', () => {
    const result = parseConversation({
      id: '19:x@thread.v2',
      lastMessage: { composetime: '2026-06-30T09:00:00Z', content: 'hey' },
    });
    expect(result?.lastMessageTime).toBe('2026-06-30T09:00:00Z');
  });
});
