/**
 * Unit tests for the new-messages polling helpers.
 *
 * Covers the opaque cursor, which conversations are worth fetching, how raw
 * chatsvc messages become compact feed items, and result ordering.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import type { ChatSummary } from './chatsvc-conversations.js';
import { httpRequest } from '../utils/http.js';
import { requireMessageAuthWithConfig } from '../utils/auth-guards.js';
import { ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import {
  encodeNewMessagesCursor,
  decodeNewMessagesCursor,
  selectConversationsToScan,
  parseNewMessage,
  orderNewMessages,
  truncateSnippet,
  getNewMessages,
  type NewMessageItem,
} from './chatsvc-new-messages.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({ requireMessageAuthWithConfig: vi.fn() }));

const mockHttp = vi.mocked(httpRequest);
const mockAuth = vi.mocked(requireMessageAuthWithConfig);
const USER_MRI = '8:orgid:me';

// ─────────────────────────────────────────────────────────────────────────────
// Cursor
// ─────────────────────────────────────────────────────────────────────────────

describe('new-message cursor', () => {
  it('round-trips the poll floor and per-conversation watermarks', () => {
    const cursor = {
      version: 1 as const,
      since: 1_700_000_000_000,
      chats: { '19:abc@thread.v2': '1700000000123' },
    };

    expect(decodeNewMessagesCursor(encodeNewMessagesCursor(cursor))).toEqual(cursor);
  });

  it('produces an opaque value that does not expose raw JSON', () => {
    const encoded = encodeNewMessagesCursor({
      version: 1,
      since: 1,
      chats: {},
    });

    expect(encoded).not.toContain('{');
    expect(encoded).not.toContain('since');
  });

  it('rejects missing, malformed and unknown-version cursors', () => {
    const unknownVersion = Buffer.from(JSON.stringify({ version: 2, since: 1, chats: {} })).toString('base64url');
    const missingSince = Buffer.from(JSON.stringify({ version: 1, chats: {} })).toString('base64url');
    const badChats = Buffer.from(JSON.stringify({ version: 1, since: 1, chats: [] })).toString('base64url');

    expect(decodeNewMessagesCursor('')).toBeNull();
    expect(decodeNewMessagesCursor('not-base64-json')).toBeNull();
    expect(decodeNewMessagesCursor(unknownVersion)).toBeNull();
    expect(decodeNewMessagesCursor(missingSince)).toBeNull();
    expect(decodeNewMessagesCursor(badChats)).toBeNull();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Conversation selection
// ─────────────────────────────────────────────────────────────────────────────

const conversation = (overrides: Partial<ChatSummary> & Pick<ChatSummary, 'conversationId'>): ChatSummary => ({
  topic: '',
  chatType: 'group',
  lastMessageTime: '2026-06-30T10:00:00Z',
  lastMessageFrom: '',
  lastMessagePreview: '',
  ...overrides,
});

describe('selectConversationsToScan', () => {
  it('only selects conversations with activity after the poll floor', () => {
    const conversations = [
      conversation({ conversationId: '19:new@thread.v2', lastMessageTime: '2026-06-30T10:00:00Z' }),
      conversation({ conversationId: '19:old@thread.v2', lastMessageTime: '2026-06-30T08:00:00Z' }),
    ];

    const result = selectConversationsToScan(
      conversations,
      null,
      new Date('2026-06-30T09:00:00Z').getTime(),
      30
    );

    expect(result.conversations.map(c => c.conversationId)).toEqual(['19:new@thread.v2']);
    expect(result.truncated).toBe(false);
  });

  it('skips conversations whose last message the cursor already covers', () => {
    const conversations = [
      conversation({ conversationId: '19:seen@thread.v2', lastMessageTime: '2026-06-30T10:00:00Z' }),
    ];

    const result = selectConversationsToScan(
      conversations,
      {
        version: 1,
        since: new Date('2026-06-30T09:00:00Z').getTime(),
        chats: { '19:seen@thread.v2': String(new Date('2026-06-30T10:00:00Z').getTime()) },
      },
      new Date('2026-06-30T09:00:00Z').getTime(),
      30
    );

    expect(result.conversations).toEqual([]);
  });

  it('takes the most recently active conversations and reports truncation', () => {
    const conversations = [
      conversation({ conversationId: '19:a@thread.v2', lastMessageTime: '2026-06-30T10:00:00Z' }),
      conversation({ conversationId: '19:b@thread.v2', lastMessageTime: '2026-06-30T09:30:00Z' }),
      conversation({ conversationId: '19:c@thread.v2', lastMessageTime: '2026-06-30T09:15:00Z' }),
    ];

    const result = selectConversationsToScan(
      conversations,
      null,
      new Date('2026-06-30T09:00:00Z').getTime(),
      2
    );

    expect(result.conversations.map(c => c.conversationId)).toEqual([
      '19:a@thread.v2',
      '19:b@thread.v2',
    ]);
    expect(result.truncated).toBe(true);
  });

  it('ignores conversations with no parsable last-message time', () => {
    const result = selectConversationsToScan(
      [conversation({ conversationId: '19:x@thread.v2', lastMessageTime: '' })],
      null,
      0,
      30
    );

    expect(result.conversations).toEqual([]);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Message parsing
// ─────────────────────────────────────────────────────────────────────────────

const context = {
  conversationId: '19:abc@thread.v2',
  topic: 'Project',
  chatType: 'group' as const,
  userMri: '8:orgid:me',
};

describe('parseNewMessage', () => {
  it('maps a message into a compact item with stripped HTML', () => {
    const item = parseNewMessage(
      {
        id: '1760000000123',
        messagetype: 'RichText/Html',
        content: '<p>Hello <b>team</b> &amp; welcome</p>',
        from: '8:orgid:alice',
        imdisplayname: 'Alice',
        originalarrivaltime: '2026-06-30T10:00:00Z',
      },
      context
    );

    expect(item).toEqual({
      conversationId: '19:abc@thread.v2',
      topic: 'Project',
      chatType: 'group',
      messageId: '1760000000123',
      senderName: 'Alice',
      sentAt: '2026-06-30T10:00:00.000Z',
      snippet: 'Hello team & welcome',
    });
  });

  it('falls back to the numeric message id for the timestamp', () => {
    const item = parseNewMessage(
      { id: '1760000000123', messagetype: 'Text', content: 'hi', from: '8:orgid:alice' },
      context
    );

    expect(item?.sentAt).toBe(new Date(1760000000123).toISOString());
  });

  it('skips system, control, thread-activity and deleted messages', () => {
    expect(parseNewMessage({ id: '1', messagetype: 'Control/Typing' }, context)).toBeNull();
    expect(parseNewMessage({ id: '1', messagetype: 'ThreadActivity/AddMember' }, context)).toBeNull();
    expect(parseNewMessage({ id: '1', messagetype: 'ThreadActivity/TopicUpdate' }, context)).toBeNull();
    expect(
      parseNewMessage(
        { id: '1', messagetype: 'Text', content: 'gone', properties: { deletetime: 'x' } },
        context
      )
    ).toBeNull();
  });

  it('skips messages without an id or message type', () => {
    expect(parseNewMessage({ messagetype: 'Text', content: 'x' }, context)).toBeNull();
    expect(parseNewMessage({ id: '1', content: 'x' }, context)).toBeNull();
  });

  it('caps the snippet at 200 characters including the ellipsis', () => {
    const item = parseNewMessage(
      { id: '1760000000123', messagetype: 'Text', content: 'a'.repeat(500), from: '8:orgid:alice' },
      context
    );

    expect(item?.snippet.length).toBe(200);
    expect(item?.snippet.endsWith('…')).toBe(true);
  });
});

describe('truncateSnippet', () => {
  it('leaves short text unchanged', () => {
    expect(truncateSnippet('hello', 200)).toBe('hello');
  });

  it('truncates to exactly the limit, ellipsis included', () => {
    expect(truncateSnippet('a'.repeat(10), 5)).toBe('aaaa…');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Ordering
// ─────────────────────────────────────────────────────────────────────────────

describe('orderNewMessages', () => {
  it('orders oldest first, then by numeric message id', () => {
    const item = (messageId: string, sentAt: string): NewMessageItem => ({
      conversationId: '19:x@thread.v2',
      topic: '',
      chatType: 'group',
      messageId,
      senderName: '',
      sentAt,
      snippet: '',
    });

    const ordered = orderNewMessages([
      item('20', '2026-06-30T10:00:02Z'),
      item('10', '2026-06-30T10:00:01Z'),
      item('11', '2026-06-30T10:00:01Z'),
    ]);

    expect(ordered.map(i => i.messageId)).toEqual(['10', '11', '20']);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Orchestration (mocked HTTP)
// ─────────────────────────────────────────────────────────────────────────────

const BASE = 'https://teams.microsoft.com';
const REGION = 'amer';
const CONVERSATION_ID = '19:abc@thread.v2';
const FLOOR = new Date('2026-06-30T09:00:00Z').getTime();

const rawConversation = (lastMessageTime: string) => ({
  id: CONVERSATION_ID,
  threadProperties: { topic: 'Project' },
  lastMessage: {
    id: String(new Date(lastMessageTime).getTime()),
    originalarrivaltime: lastMessageTime,
    imdisplayname: 'Alice',
    content: 'latest',
  },
});

const httpOk = <T>(data: T) => ok({ status: 200, headers: new Headers(), data });

beforeEach(() => {
  vi.clearAllMocks();
  mockAuth.mockReturnValue(
    ok({
      auth: { skypeToken: 'skype', authToken: 'bearer', userMri: USER_MRI },
      region: REGION,
      baseUrl: BASE,
    })
  );
});

describe('getNewMessages', () => {
  it('costs one request when no conversation has new activity', async () => {
    mockHttp.mockResolvedValue(
      httpOk({
        conversations: [rawConversation('2026-06-30T08:30:00Z')],
      })
    );

    const result = await getNewMessages({
      after: encodeNewMessagesCursor({ version: 1, since: FLOOR, chats: {} }),
      now: new Date('2026-06-30T10:00:00Z').getTime(),
    });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(mockHttp).toHaveBeenCalledTimes(1);
    expect(result.value.items).toEqual([]);
  });

  it('fetches only active conversations and returns a compact item', async () => {
    mockHttp.mockImplementation(async url => {
      if (url.includes('/conversations?')) {
        return httpOk({ conversations: [rawConversation('2026-06-30T09:30:00Z')] });
      }
      return httpOk({
        messages: [
          {
            id: String(new Date('2026-06-30T09:30:00Z').getTime()),
            messagetype: 'RichText/Html',
            content: '<p>Hello from Alice</p>',
            from: '8:orgid:alice',
            imdisplayname: 'Alice',
            originalarrivaltime: '2026-06-30T09:30:00Z',
          },
        ],
      });
    });

    const result = await getNewMessages({
      after: encodeNewMessagesCursor({ version: 1, since: FLOOR, chats: {} }),
      now: new Date('2026-06-30T10:00:00Z').getTime(),
    });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.items).toEqual([
      {
        conversationId: CONVERSATION_ID,
        topic: 'Project',
        chatType: 'group',
        messageId: String(new Date('2026-06-30T09:30:00Z').getTime()),
        senderName: 'Alice',
        sentAt: '2026-06-30T09:30:00.000Z',
        snippet: 'Hello from Alice',
      },
    ]);
    expect(result.value.truncated).toBe(false);
    expect(result.value.errors).toEqual([]);

    const cursor = decodeNewMessagesCursor(result.value.nextAfter);
    expect(cursor?.chats[CONVERSATION_ID]).toBe(String(new Date('2026-06-30T09:30:00Z').getTime()));
  });

  it('excludes own messages but still advances the watermark past them', async () => {
    mockHttp.mockImplementation(async url => {
      if (url.includes('/conversations?')) {
        return httpOk({ conversations: [rawConversation('2026-06-30T09:40:00Z')] });
      }
      return httpOk({
        messages: [
          {
            id: String(new Date('2026-06-30T09:30:00Z').getTime()),
            messagetype: 'Text',
            content: 'from Alice',
            from: '8:orgid:alice',
            imdisplayname: 'Alice',
            originalarrivaltime: '2026-06-30T09:30:00Z',
          },
          {
            id: String(new Date('2026-06-30T09:40:00Z').getTime()),
            messagetype: 'Text',
            content: 'my reply',
            from: USER_MRI,
            imdisplayname: 'Me',
            originalarrivaltime: '2026-06-30T09:40:00Z',
          },
        ],
      });
    });

    const result = await getNewMessages({
      after: encodeNewMessagesCursor({ version: 1, since: FLOOR, chats: {} }),
      now: new Date('2026-06-30T10:00:00Z').getTime(),
    });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.items.map(item => item.snippet)).toEqual(['from Alice']);
    const cursor = decodeNewMessagesCursor(result.value.nextAfter);
    expect(cursor?.chats[CONVERSATION_ID]).toBe(String(new Date('2026-06-30T09:40:00Z').getTime()));
  });

  it('reports per-conversation failures without advancing the poll floor', async () => {
    mockHttp.mockImplementation(async url => {
      if (url.includes('/conversations?')) {
        return httpOk({ conversations: [rawConversation('2026-06-30T09:30:00Z')] });
      }
      return err(createError(ErrorCode.API_ERROR, 'thread unavailable'));
    });

    const result = await getNewMessages({
      after: encodeNewMessagesCursor({ version: 1, since: FLOOR, chats: {} }),
      now: new Date('2026-06-30T10:00:00Z').getTime(),
    });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.items).toEqual([]);
    expect(result.value.errors).toEqual([
      { conversationId: CONVERSATION_ID, code: ErrorCode.API_ERROR, message: 'thread unavailable' },
    ]);
    expect(decodeNewMessagesCursor(result.value.nextAfter)?.since).toBe(FLOOR);
  });

  it('marks truncation and keeps the floor when maxChats hides activity', async () => {
    const secondId = '19:second@thread.v2';
    mockHttp.mockImplementation(async url => {
      if (url.includes('/conversations?')) {
        return httpOk({
          conversations: [
            rawConversation('2026-06-30T09:50:00Z'),
            {
              id: secondId,
              lastMessage: {
                id: String(new Date('2026-06-30T09:40:00Z').getTime()),
                originalarrivaltime: '2026-06-30T09:40:00Z',
                content: 'second',
              },
            },
          ],
        });
      }
      return httpOk({ messages: [] });
    });

    const result = await getNewMessages({
      after: encodeNewMessagesCursor({ version: 1, since: FLOOR, chats: {} }),
      maxChats: 1,
      now: new Date('2026-06-30T10:00:00Z').getTime(),
    });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.truncated).toBe(true);
    expect(decodeNewMessagesCursor(result.value.nextAfter)?.since).toBe(FLOOR);
  });

  it('returns the newest page and no start bound on the first run', async () => {
    mockHttp.mockImplementation(async url => {
      if (url.includes('/conversations?')) {
        return httpOk({ conversations: [rawConversation('2026-06-30T09:00:00Z')] });
      }
      return httpOk({ messages: [] });
    });

    await getNewMessages({ now: new Date('2026-06-30T10:00:00Z').getTime() });

    const messageCalls = mockHttp.mock.calls
      .map(([url]) => url as string)
      .filter(url => url.includes('/messages?'));
    expect(messageCalls).toHaveLength(1);
    expect(messageCalls[0]).not.toContain('startTime=');
  });

  it('does not flag first-run history as truncated', async () => {
    mockHttp.mockImplementation(async url => {
      if (url.includes('/conversations?')) {
        return httpOk({ conversations: [rawConversation('2026-06-30T09:00:00Z')] });
      }
      return httpOk({
        messages: Array.from({ length: 50 }, (_, index) => ({
          id: String(new Date('2026-06-30T09:00:00Z').getTime() + index),
          messagetype: 'Text',
          content: `message ${index}`,
          from: '8:orgid:alice',
        })),
      });
    });

    const result = await getNewMessages({ now: new Date('2026-06-30T10:00:00Z').getTime() });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.items).toHaveLength(50);
    expect(result.value.truncated).toBe(false);
  });

  it('advances past a channel whose latest activity is only a thread reply', async () => {
    mockHttp.mockImplementation(async url => {
      if (url.includes('/conversations?')) {
        return httpOk({ conversations: [rawConversation('2026-06-30T09:30:00Z')] });
      }
      // The channel feed has no top-level post at the new activity time.
      return httpOk({ messages: [] });
    });

    const result = await getNewMessages({
      after: encodeNewMessagesCursor({ version: 1, since: FLOOR, chats: {} }),
      now: new Date('2026-06-30T10:00:00Z').getTime(),
    });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const cursor = decodeNewMessagesCursor(result.value.nextAfter);
    expect(cursor?.chats[CONVERSATION_ID]).toBe(String(new Date('2026-06-30T09:30:00Z').getTime()));
  });

  it('keeps a clock-skew overlap when advancing the poll floor', async () => {
    const now = new Date('2026-06-30T10:00:00Z').getTime();
    mockHttp.mockImplementation(async url => {
      if (url.includes('/conversations?')) {
        return httpOk({ conversations: [rawConversation('2026-06-30T09:30:00Z')] });
      }
      return httpOk({ messages: [] });
    });

    const result = await getNewMessages({
      after: encodeNewMessagesCursor({ version: 1, since: FLOOR, chats: {} }),
      now,
    });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(decodeNewMessagesCursor(result.value.nextAfter)?.since).toBe(now - 60_000);
  });

  it('rejects a malformed cursor', async () => {
    const result = await getNewMessages({ after: 'not-a-cursor' });

    expect(result).toMatchObject({ ok: false, error: { code: ErrorCode.INVALID_INPUT } });
    expect(mockHttp).not.toHaveBeenCalled();
  });
});

describe('resuming successive polls', () => {
  const ids = [FLOOR + 3, FLOOR + 2, FLOOR + 1, FLOOR];
  const knownChat = encodeNewMessagesCursor({ version: 1, since: FLOOR + 10, chats: { [CONVERSATION_ID]: String(FLOOR) } });
  const now = FLOOR + 120_000;

  it.each([false, true])('drains every page once with an inclusive start time (first request fails: %s)', async failFirst => {
    let fail = failFirst;
    let newest = ids[0];
    mockHttp.mockImplementation(async url => {
      if (url.includes('/conversations?')) {
        return httpOk({ conversations: [rawConversation(new Date(newest).toISOString())] });
      }
      if (fail) {
        fail = false;
        return err(createError(ErrorCode.API_ERROR, 'temporary failure'));
      }
      const params = new URL(url).searchParams;
      const offset = Number(params.get('syncState'));
      if (offset) newest = FLOOR + 4;
      const messages = [{ id: String(ids[offset]), messagetype: 'Text' }];
      const backwardLink = BASE + '/messages?syncState=' + (offset + 1);
      return httpOk({ messages, _metadata: { backwardLink } });
    });

    let after = knownChat;
    const received = [];
    for (let i = 0; i < 5; i++) {
      const result = await getNewMessages({ after, limitPerChat: 1, now });
      if (!result.ok) throw Error(result.error.message);
      received.push(...result.value.items.map(item => Number(item.messageId)));
      after = result.value.nextAfter;
      if (!result.value.truncated && !result.value.errors.length) break;
    }
    expect(received).toEqual(ids.slice(0, 3));
    expect(decodeNewMessagesCursor(after)?.chats[CONVERSATION_ID]).toBe(String(ids[0]));
    // The new arrival was never part of the older pages.
    mockHttp.mockImplementation(async url => url.includes('/conversations?')
      ? httpOk({ conversations: [rawConversation(new Date(newest).toISOString())] })
      : httpOk({ messages: [{ id: String(newest), messagetype: 'Text' }] }));
    const next = await getNewMessages({ after, now });
    if (!next.ok) throw Error(next.error.message);
    expect(next.value.items.map(i => Number(i.messageId))).toEqual([newest]);
  });

  it('drains a second conversation-list page and picks up activity during the sweep', async () => {
    let later = false;
    const chat = (id: number, time: number) => ({
      ...rawConversation(new Date(time).toISOString()), id: 'chat-' + id,
    });
    mockHttp.mockImplementation(async url => {
      if (url.includes('/conversations?')) {
        if (new URL(url).searchParams.get('syncState') === 'second') {
          return httpOk({ conversations: [chat(201, FLOOR)] });
        }
        return httpOk({
          conversations: [
            ...(later ? [chat(202, now + 1)] : []),
            ...Array.from({ length: later ? 199 : 200 }, (_, i) => chat(i, FLOOR + 1)),
          ],
          _metadata: { backwardLink: BASE + '/conversations?syncState=second' },
        });
      }
      const id = url.includes('/chat-202/') ? now + 1 : url.includes('/chat-201/') ? FLOOR : FLOOR + 1;
      return httpOk({ messages: [{ id: String(id), messagetype: 'Text' }] });
    });
    const first = await getNewMessages({ maxChats: 200, now });
    if (!first.ok) throw Error(first.error.message);
    expect(first.value.items).toHaveLength(200);
    expect(first.value.truncated).toBe(true);

    later = true;
    const second = await getNewMessages({ after: first.value.nextAfter, now: now + 120_000 });
    if (!second.ok) throw Error(second.error.message);
    expect(second.value.items.map(i => i.conversationId)).toEqual(['chat-201']);
    expect(second.value.truncated).toBe(false);

    const next = await getNewMessages({ after: second.value.nextAfter, now: now + 120_001 });
    if (!next.ok) throw Error(next.error.message);
    expect(next.value.items.map(i => i.conversationId)).toEqual(['chat-202']);
  });
});
