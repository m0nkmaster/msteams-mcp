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
} from './chatsvc-messaging.js';

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

  it('keeps inline span itemids aligned with the mentions array for multiple mentions', () => {
    // Regression: the inline <span itemid> was numbered in reverse
    // placeholder-processing order while the mentions array (and
    // buildMentionsProperty, which keys by array index) stayed source-ordered.
    // That mismatch reversed the mention→MRI mapping: one mention was fine, two
    // swapped, three reversed (middle fixed).
    const { html, mentions } = parseContentWithMentionsAndLinks(
      '@[Noah](8:orgid:noah) @[Kubra](8:orgid:kubra) @[Atakan](8:orgid:atakan)'
    );

    expect(mentions.map((m) => m.mri)).toEqual([
      '8:orgid:noah',
      '8:orgid:kubra',
      '8:orgid:atakan',
    ]);

    // The itemid on each name's span must equal that mention's index in the array.
    const itemIdFor = (name: string): number => {
      const match = html.match(new RegExp(`itemid="(\\d+)"[^>]*>${name}<`));
      if (!match) throw new Error(`no mention span found for ${name}`);
      return Number(match[1]);
    };
    expect(itemIdFor('Noah')).toBe(0);
    expect(itemIdFor('Kubra')).toBe(1);
    expect(itemIdFor('Atakan')).toBe(2);

    // And the property built from that array resolves each span's itemid back to
    // the correct MRI (this is what Teams uses to render the chip / notify).
    const props = JSON.parse(buildMentionsProperty(mentions));
    expect(props[itemIdFor('Noah')].mri).toBe('8:orgid:noah');
    expect(props[itemIdFor('Kubra')].mri).toBe('8:orgid:kubra');
    expect(props[itemIdFor('Atakan')].mri).toBe('8:orgid:atakan');
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
});
