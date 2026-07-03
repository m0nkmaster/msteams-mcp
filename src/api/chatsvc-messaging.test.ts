/**
 * Unit tests for mention and content parsing in chatsvc-messaging.
 *
 * Tests the tag vs person mention differentiation, HTML generation,
 * and mentions property serialisation.
 */

import { describe, it, expect } from 'vitest';
import {
  buildMentionHtml,
  buildMentionsProperty,
  parseContentWithMentionsAndLinks,
  resolveMessageContent,
  parseScheduleTime,
  buildScheduledDraftBody,
  extractOneOnOneMemberIds,
  buildOneOnOneThreadBody,
  clampWaitParams,
  selectNewMessages,
  resolveWaitBaseline,
  generateClientMessageId,
} from './chatsvc-messaging.js';
import { MAX_WAIT_SECONDS } from '../constants.js';

// ─────────────────────────────────────────────────────────────────────────────
// buildMentionHtml
// ─────────────────────────────────────────────────────────────────────────────

describe('buildMentionHtml', () => {
  it('wraps person mentions in readonly + span', () => {
    const html = buildMentionHtml('Alice', 0, '8:orgid:abc-123');
    expect(html).toContain('<readonly');
    expect(html).toContain('skipProofing');
    expect(html).toContain('<span');
    expect(html).toContain('itemid="0"');
    expect(html).toContain('Alice');
  });

  it('uses span-only format for tag mentions', () => {
    const html = buildMentionHtml('engineering', 0, 'tag:txk8gOnia');
    expect(html).not.toContain('<readonly');
    expect(html).toContain('<span');
    expect(html).toContain('itemid="0"');
    expect(html).toContain('engineering');
  });

  it('escapes HTML characters in display name', () => {
    const html = buildMentionHtml('O\'Brien & Co <team>', 1, 'tag:abc');
    expect(html).toContain('&amp;');
    expect(html).toContain('&lt;');
    expect(html).toContain('&gt;');
    expect(html).not.toContain('& ');
    expect(html).not.toContain('<team>');
  });

  it('uses correct itemId for sequential mentions', () => {
    const first = buildMentionHtml('Tag1', 0, 'tag:a');
    const second = buildMentionHtml('Person', 1, '8:orgid:xyz');
    expect(first).toContain('itemid="0"');
    expect(second).toContain('itemid="1"');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// buildMentionsProperty
// ─────────────────────────────────────────────────────────────────────────────

describe('buildMentionsProperty', () => {
  it('sets mentionType "person" for regular MRIs', () => {
    const result = JSON.parse(buildMentionsProperty([
      { mri: '8:orgid:abc-123', displayName: 'Alice' },
    ]));
    expect(result).toHaveLength(1);
    expect(result[0].mentionType).toBe('person');
    expect(result[0].mri).toBe('8:orgid:abc-123');
  });

  it('sets mentionType "tag" and strips prefix for tag MRIs', () => {
    const result = JSON.parse(buildMentionsProperty([
      { mri: 'tag:txk8gOnia', displayName: 'engineering' },
    ]));
    expect(result).toHaveLength(1);
    expect(result[0].mentionType).toBe('tag');
    expect(result[0].mri).toBe('txk8gOnia');
    expect(result[0].displayName).toBe('engineering');
  });

  it('handles mixed person and tag mentions', () => {
    const result = JSON.parse(buildMentionsProperty([
      { mri: '8:orgid:abc', displayName: 'Alice' },
      { mri: 'tag:xyz', displayName: 'my-tag' },
    ]));
    expect(result).toHaveLength(2);
    expect(result[0].mentionType).toBe('person');
    expect(result[0].mri).toBe('8:orgid:abc');
    expect(result[1].mentionType).toBe('tag');
    expect(result[1].mri).toBe('xyz');
  });

  it('assigns sequential itemids', () => {
    const result = JSON.parse(buildMentionsProperty([
      { mri: '8:orgid:a', displayName: 'A' },
      { mri: 'tag:b', displayName: 'B' },
      { mri: '8:orgid:c', displayName: 'C' },
    ]));
    expect(result[0].itemid).toBe('0');
    expect(result[1].itemid).toBe('1');
    expect(result[2].itemid).toBe('2');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// parseContentWithMentionsAndLinks - tag mentions
// ─────────────────────────────────────────────────────────────────────────────

describe('parseContentWithMentionsAndLinks', () => {
  it('parses a person mention', () => {
    const { html, mentions } = parseContentWithMentionsAndLinks(
      'Hey @[Alice](8:orgid:abc), check this'
    );
    expect(mentions).toHaveLength(1);
    expect(mentions[0].mri).toBe('8:orgid:abc');
    expect(html).toContain('<readonly');
    expect(html).toContain('Alice');
  });

  it('parses a tag mention with span-only HTML', () => {
    const { html, mentions } = parseContentWithMentionsAndLinks(
      'Hey @[engineering](tag:txk8gOnia), please review'
    );
    expect(mentions).toHaveLength(1);
    expect(mentions[0].mri).toBe('tag:txk8gOnia');
    expect(mentions[0].displayName).toBe('engineering');
    expect(html).not.toContain('<readonly');
    expect(html).toContain('<span');
    expect(html).toContain('engineering');
  });

  it('handles mixed person and tag mentions in order', () => {
    const { html, mentions } = parseContentWithMentionsAndLinks(
      '@[Alice](8:orgid:abc) and @[my-tag](tag:xyz123) - thoughts?'
    );
    expect(mentions).toHaveLength(2);
    expect(mentions[0].mri).toBe('8:orgid:abc');
    expect(mentions[1].mri).toBe('tag:xyz123');
    // Person mention should have readonly wrapper
    expect(html).toContain('<readonly');
    // Both names present
    expect(html).toContain('Alice');
    expect(html).toContain('my-tag');
  });

  it('returns markdown-converted HTML when no mentions or links', () => {
    const { html, mentions } = parseContentWithMentionsAndLinks('Just plain text');
    expect(mentions).toHaveLength(0);
    expect(html).toContain('Just plain text');
  });

  it('handles links alongside tag mentions', () => {
    const { html, mentions } = parseContentWithMentionsAndLinks(
      '@[my-tag](tag:abc) see [docs](https://example.com)'
    );
    expect(mentions).toHaveLength(1);
    expect(mentions[0].mri).toBe('tag:abc');
    expect(html).toContain('<a href="https://example.com">docs</a>');
    expect(html).toContain('my-tag');
  });

  it('renders a mailto link', () => {
    const { html } = parseContentWithMentionsAndLinks('Email [me](mailto:me@example.com) now');
    expect(html).toContain('<a href="mailto:me@example.com">me</a>');
  });

  it('never emits a live href for a javascript: link', () => {
    const { html } = parseContentWithMentionsAndLinks('[click](javascript:alert(1))');
    expect(html).not.toContain('href="javascript:');
  });
});

describe('resolveMessageContent', () => {
  it('defaults to markdown conversion (backward compatible)', () => {
    const r = resolveMessageContent('This is **bold**');
    expect(r.messagetype).toBe('RichText/Html');
    expect(r.content).toContain('<b>bold</b>');
    expect(r.mentions).toHaveLength(0);
  });

  it('markdown mode extracts mentions', () => {
    const r = resolveMessageContent('Hi @[Alice](8:orgid:abc)', 'markdown');
    expect(r.messagetype).toBe('RichText/Html');
    expect(r.mentions).toHaveLength(1);
    expect(r.mentions[0].mri).toBe('8:orgid:abc');
  });

  it('text mode sends content verbatim with no conversion or mentions', () => {
    const raw = 'Cost is 5*3*2 and use a_b_c @[A](8:orgid:x)';
    const r = resolveMessageContent(raw, 'text');
    expect(r.messagetype).toBe('Text');
    expect(r.content).toBe(raw);
    expect(r.mentions).toHaveLength(0);
  });

  it('html mode passes content through as RichText/Html without conversion', () => {
    const r = resolveMessageContent('<p>hi <b>there</b></p>', 'html');
    expect(r.messagetype).toBe('RichText/Html');
    expect(r.content).toBe('<p>hi <b>there</b></p>');
    expect(r.mentions).toHaveLength(0);
  });

  it('auto mode sends plain single-line text verbatim as Text', () => {
    const r = resolveMessageContent('just plain text', 'auto');
    expect(r.messagetype).toBe('Text');
    expect(r.content).toBe('just plain text');
  });

  it('auto mode converts when markdown formatting is present', () => {
    const r = resolveMessageContent('This is **bold**', 'auto');
    expect(r.messagetype).toBe('RichText/Html');
    expect(r.content).toContain('<b>bold</b>');
  });

  it('auto mode converts when a mention is present even without markdown', () => {
    const r = resolveMessageContent('hi @[Alice](8:orgid:abc)', 'auto');
    expect(r.messagetype).toBe('RichText/Html');
    expect(r.mentions).toHaveLength(1);
  });
});

describe('parseScheduleTime', () => {
  const NOW = Date.parse('2026-01-01T00:00:00Z');

  it('accepts an ISO 8601 UTC datetime in the future', () => {
    const r = parseScheduleTime('2099-12-31T09:00:00Z', NOW);
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.value.iso).toBe('2099-12-31T09:00:00.000Z');
      expect(r.value.epochMs).toBe(Date.parse('2099-12-31T09:00:00Z'));
    }
  });

  it('converts a datetime with offset to UTC', () => {
    const r = parseScheduleTime('2099-12-31T09:00:00+05:00', NOW);
    expect(r.ok).toBe(true);
    if (r.ok) expect(r.value.iso).toBe('2099-12-31T04:00:00.000Z');
  });

  it('treats a timezone-less datetime as UTC', () => {
    const r = parseScheduleTime('2099-12-31T09:00:00', NOW);
    expect(r.ok).toBe(true);
    if (r.ok) expect(r.value.iso).toBe('2099-12-31T09:00:00.000Z');
  });

  it('accepts a space separator without seconds (as UTC)', () => {
    const r = parseScheduleTime('2099-12-31 09:00', NOW);
    expect(r.ok).toBe(true);
    if (r.ok) expect(r.value.iso).toBe('2099-12-31T09:00:00.000Z');
  });

  it('rejects a past datetime', () => {
    const r = parseScheduleTime('2020-01-01T00:00:00Z', NOW);
    expect(r.ok).toBe(false);
  });

  it('rejects an unparseable datetime', () => {
    expect(parseScheduleTime('not-a-date', NOW).ok).toBe(false);
    expect(parseScheduleTime('', NOW).ok).toBe(false);
  });
});

describe('buildScheduledDraftBody', () => {
  it('builds a ScheduledDraft body with sendAt in epoch millis', () => {
    const epochMs = Date.parse('2099-12-31T09:00:00Z');
    const body = buildScheduledDraftBody({
      conversationId: '19:abc@thread.v2',
      content: 'Hello later',
      messagetype: 'Text',
      epochMs,
      userMri: '8:orgid:me',
      isoNow: '2026-01-01T00:00:00.000Z',
      conversationLink: 'https://teams.microsoft.com/api/chatsvc/amer/v1/users/ME/conversations/19:abc@thread.v2',
      clientMessageId: '12345',
    });
    expect(body.draftType).toBe('ScheduledDraft');
    expect((body.draftDetails as Record<string, unknown>).sendAt).toBe(epochMs);
    expect(body.innerThreadId).toBe('19:abc@thread.v2');
    const message = body.message as Record<string, unknown>;
    expect(message.content).toBe('Hello later');
    expect(message.messagetype).toBe('Text');
    expect(message.from).toBe('8:orgid:me');
    expect((message.draftDetails as Record<string, unknown>).sendAt).toBe(epochMs);
  });
});

describe('extractOneOnOneMemberIds', () => {
  it('extracts both member object IDs from a 1:1 conversation ID', () => {
    expect(extractOneOnOneMemberIds('19:aaa_bbb@unq.gbl.spaces')).toEqual(['aaa', 'bbb']);
  });

  it('returns null for non 1:1 conversation IDs', () => {
    expect(extractOneOnOneMemberIds('19:xyz@thread.tacv2')).toBeNull();
    expect(extractOneOnOneMemberIds('19:xyz@thread.v2')).toBeNull();
  });

  it('returns null for a malformed 1:1 ID missing a member', () => {
    expect(extractOneOnOneMemberIds('19:aaa_@unq.gbl.spaces')).toBeNull();
    expect(extractOneOnOneMemberIds('19:_bbb@unq.gbl.spaces')).toBeNull();
  });
});

describe('buildOneOnOneThreadBody', () => {
  it('builds a unique-roster thread body with both members as Admins', () => {
    const body = buildOneOnOneThreadBody(['aaa', 'bbb']);
    const members = body.members as Array<{ id: string; role: string }>;
    expect(members).toHaveLength(2);
    expect(members[0]).toEqual({ id: '8:orgid:aaa', role: 'Admin' });
    expect(members[1]).toEqual({ id: '8:orgid:bbb', role: 'Admin' });
    const props = body.properties as Record<string, string>;
    expect(props.threadType).toBe('chat');
    expect(props.uniquerosterthread).toBe('true');
    expect(props.fixedRoster).toBe('true');
  });
});

describe('clampWaitParams', () => {
  it('leaves in-range values unchanged', () => {
    expect(clampWaitParams(60, 5)).toEqual({ maxWait: 60, interval: 5 });
  });

  it('floors maxWait and interval to at least 1 second', () => {
    expect(clampWaitParams(0, 0)).toEqual({ maxWait: 1, interval: 1 });
  });

  it('caps maxWait at MAX_WAIT_SECONDS', () => {
    expect(clampWaitParams(99999, 10)).toEqual({ maxWait: MAX_WAIT_SECONDS, interval: 10 });
  });

  it('never lets the interval exceed the clamped maxWait', () => {
    expect(clampWaitParams(30, 600)).toEqual({ maxWait: 30, interval: 30 });
  });
});

describe('selectNewMessages', () => {
  const msgs = [
    { id: '100', isFromMe: false },
    { id: '105', isFromMe: true },
    { id: '110', isFromMe: false },
    { id: 'abc', isFromMe: false }, // non-numeric id must be ignored
  ];

  it('returns only messages strictly newer than afterId, excluding own by default', () => {
    const result = selectNewMessages(msgs, 100, false);
    expect(result.map(m => m.id)).toEqual(['110']);
  });

  it('includes own messages when includeSelf is true, sorted ascending', () => {
    const result = selectNewMessages(msgs, 100, true);
    expect(result.map(m => m.id)).toEqual(['105', '110']);
  });

  it('returns empty when nothing is newer than afterId', () => {
    expect(selectNewMessages(msgs, 110, true)).toEqual([]);
  });
});

describe('generateClientMessageId', () => {
  it('returns a numeric string (Teams rejects non-numeric client message IDs)', () => {
    // Regression: a UUID (crypto.randomUUID) is rejected by Teams with
    // "StoreInvalidInput - ClientMessageId must be a number in string format".
    const id = generateClientMessageId();
    expect(id).toMatch(/^\d+$/);
    expect(id.length).toBeGreaterThan(0);
  });

  it('produces distinct IDs across rapid calls', () => {
    const ids = new Set(Array.from({ length: 50 }, () => generateClientMessageId()));
    expect(ids.size).toBeGreaterThan(1);
  });
});

describe('resolveWaitBaseline', () => {
  it('prefers the caller\'s own most recent message', () => {
    const baseline = resolveWaitBaseline([
      { id: '100', isFromMe: false },
      { id: '90', isFromMe: true },
      { id: '110', isFromMe: false },
    ]);
    expect(baseline).toBe(90);
  });

  it('falls back to the most recent message overall when the caller has not posted', () => {
    const baseline = resolveWaitBaseline([
      { id: '100', isFromMe: false },
      { id: '110', isFromMe: false },
    ]);
    expect(baseline).toBe(110);
  });

  it('returns 0 for an empty thread', () => {
    expect(resolveWaitBaseline([])).toBe(0);
  });
});
