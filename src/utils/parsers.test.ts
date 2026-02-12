/**
 * Unit tests for parsing functions.
 * 
 * Tests outcomes, not implementations - verify that given inputs
 * produce expected outputs regardless of internal logic.
 */

import { describe, it, expect } from 'vitest';
import {
  stripHtml,
  extractLinks,
  buildMessageLink,
  getConversationType,
  extractMessageTimestamp,
  parsePersonSuggestion,
  parseV2Result,
  parseJwtProfile,
  calculateTokenStatus,
  parseSearchResults,
  parsePeopleResults,
  extractObjectId,
  buildOneOnOneConversationId,
  decodeBase64Guid,
  extractActivityTimestamp,
  markdownToTeamsHtml,
  formatTranscriptText,
  parseChannelSuggestion,
  parseChannelResults,
  filterChannelsByName,
  parseTeamsList,
  parseVirtualConversationMessage,
  hasMarkdownFormatting,
} from './parsers.js';
import {
  searchResultItem,
  searchResultWithHtml,
  searchResultMinimal,
  searchResultTooShort,
  searchResultThreadReply,
  searchEntitySetsResponse,
  personSuggestion,
  personMinimal,
  personWithBase64Id,
  peopleGroupsResponse,
  jwtPayloadFull,
  jwtPayloadMinimal,
  jwtPayloadCommaName,
  jwtPayloadSpaceName,
  sourceWithMessageId,
  sourceWithConvIdMessageId,
} from '../__fixtures__/api-responses.js';

describe('stripHtml', () => {
  it('removes HTML tags', () => {
    expect(stripHtml('<p>Hello</p>')).toBe('Hello');
    expect(stripHtml('<div><strong>Bold</strong> text</div>')).toBe('Bold text');
  });

  it('decodes HTML entities', () => {
    expect(stripHtml('Tom &amp; Jerry')).toBe('Tom & Jerry');
    expect(stripHtml('1 &lt; 2 &gt; 0')).toBe('1 < 2 > 0');
    expect(stripHtml('&quot;quoted&quot;')).toBe('"quoted"');
    expect(stripHtml("it&#39;s")).toBe("it's");
    expect(stripHtml('non&nbsp;breaking')).toBe('non breaking');
  });

  it('collapses whitespace', () => {
    expect(stripHtml('hello    world')).toBe('hello world');
    expect(stripHtml('  trimmed  ')).toBe('trimmed');
    expect(stripHtml('line\n\nbreak')).toBe('line break');
  });

  it('handles complex HTML', () => {
    const html = '<p>Meeting <strong>notes</strong> from &amp; yesterday&apos;s call</p><br/><div>Action items:</div>';
    expect(stripHtml(html)).toBe("Meeting notes from & yesterday's call Action items:");
  });

  it('returns empty string for empty input', () => {
    expect(stripHtml('')).toBe('');
  });
});

describe('extractLinks', () => {
  it('extracts simple links', () => {
    const html = 'Check out <a href="https://example.com">this link</a> here';
    expect(extractLinks(html)).toEqual([
      { url: 'https://example.com', text: 'this link' }
    ]);
  });

  it('extracts multiple links', () => {
    const html = '<a href="https://a.com">A</a> and <a href="https://b.com">B</a>';
    expect(extractLinks(html)).toEqual([
      { url: 'https://a.com', text: 'A' },
      { url: 'https://b.com', text: 'B' }
    ]);
  });

  it('strips nested HTML from link text', () => {
    const html = '<a href="https://example.com"><strong>Bold</strong> link</a>';
    expect(extractLinks(html)).toEqual([
      { url: 'https://example.com', text: 'Bold link' }
    ]);
  });

  it('uses URL as text when link text is empty', () => {
    const html = '<a href="https://example.com"></a>';
    expect(extractLinks(html)).toEqual([
      { url: 'https://example.com', text: 'https://example.com' }
    ]);
  });

  it('ignores javascript: links', () => {
    const html = '<a href="javascript:void(0)">Click</a>';
    expect(extractLinks(html)).toEqual([]);
  });

  it('handles links with extra attributes', () => {
    const html = '<a class="link" href="https://example.com" target="_blank">Link</a>';
    expect(extractLinks(html)).toEqual([
      { url: 'https://example.com', text: 'Link' }
    ]);
  });

  it('returns empty array when no links', () => {
    expect(extractLinks('No links here')).toEqual([]);
    expect(extractLinks('')).toEqual([]);
  });
});

describe('getConversationType', () => {
  it('identifies channel conversations', () => {
    expect(getConversationType('19:abc@thread.tacv2')).toBe('channel');
    expect(getConversationType('19:QsLXSoyGdLTIChUa-elhfgq_VyIauBGVMBk3-7orc1w1@thread.tacv2')).toBe('channel');
  });

  it('identifies meeting conversations', () => {
    expect(getConversationType('19:meeting_OWVkMDgzYWMtOGQyNi00NjQ0@thread.v2')).toBe('meeting');
    expect(getConversationType('19:meeting_abc123@thread.v2')).toBe('meeting');
  });

  it('identifies 1:1 chat conversations', () => {
    expect(getConversationType('19:ab76f827-27e2-4c67-a765-f1a53145fa24_b71f4d0f-ed13-4f3e-abdf-037e146be579@unq.gbl.spaces')).toBe('chat');
  });

  it('identifies group chat conversations', () => {
    // Group chats use @thread.v2 but don't have meeting_ prefix
    expect(getConversationType('19:abc123@thread.v2')).toBe('chat');
  });
});

describe('buildMessageLink', () => {
  it('builds channel link with parentMessageId and createdTime', () => {
    // Channel links always include parentMessageId (defaults to messageId for top-level posts)
    const link = buildMessageLink('19:abc@thread.tacv2', '1705760000000');
    expect(link).toContain('parentMessageId=1705760000000');
    expect(link).toContain('createdTime=1705760000000');
    expect(link).not.toContain('context');
  });

  it('builds chat link with context parameter', () => {
    const link = buildMessageLink('19:guid1_guid2@unq.gbl.spaces', '1705760000000');
    expect(link).toContain('context=%7B%22contextType%22%3A%22chat%22%7D');
  });

  it('builds meeting link with context parameter', () => {
    const link = buildMessageLink('19:meeting_abc@thread.v2', 1705760000000);
    expect(link).toContain('context=%7B%22contextType%22%3A%22chat%22%7D');
  });

  it('builds group chat link with context parameter', () => {
    const link = buildMessageLink('19:abc@thread.v2', 1705760000000);
    expect(link).toContain('context=%7B%22contextType%22%3A%22chat%22%7D');
  });

  it('builds channel thread reply link with parentMessageId', () => {
    // Thread reply: message timestamp differs from parent
    const link = buildMessageLink('19:abc@thread.tacv2', '1705770000000', '1705760000000');
    expect(link).toContain('parentMessageId=1705760000000');
    expect(link).toContain('createdTime=1705770000000');
  });

  it('includes parentMessageId for top-level channel posts (defaults to messageId)', () => {
    // Per MS docs, parentMessageId is always required for channel links
    const link = buildMessageLink('19:abc@thread.tacv2', '1705760000000');
    expect(link).toContain('parentMessageId=1705760000000');
  });

  it('encodes special characters in conversation ID', () => {
    const link = buildMessageLink('19:special@thread.tacv2', '123');
    expect(link).toContain('19%3Aspecial%40thread.tacv2');
  });

  it('builds channel link with tenantId and groupId via options object', () => {
    const link = buildMessageLink({
      conversationId: '19:abc@thread.tacv2',
      messageId: '1705760000000',
      tenantId: '0d9b645f-597b-41f0-a2a3-ef103fbd91bb',
      groupId: '3606f714-ec2e-41b3-9ad1-6afb331bd35d',
    });
    expect(link).toContain('tenantId=0d9b645f-597b-41f0-a2a3-ef103fbd91bb');
    expect(link).toContain('groupId=3606f714-ec2e-41b3-9ad1-6afb331bd35d');
    expect(link).toContain('parentMessageId=1705760000000');
    expect(link).toContain('createdTime=1705760000000');
  });

  it('builds chat link with tenantId via options object', () => {
    const link = buildMessageLink({
      conversationId: '19:guid1_guid2@unq.gbl.spaces',
      messageId: '1705760000000',
      tenantId: '0d9b645f-597b-41f0-a2a3-ef103fbd91bb',
    });
    expect(link).toContain('tenantId=0d9b645f-597b-41f0-a2a3-ef103fbd91bb');
    expect(link).toContain('context=%7B%22contextType%22%3A%22chat%22%7D');
  });

  it('uses custom teamsBaseUrl for GCC support', () => {
    const link = buildMessageLink({
      conversationId: '19:guid1_guid2@unq.gbl.spaces',
      messageId: '1705760000000',
      teamsBaseUrl: 'https://teams.microsoft.us',
    });
    expect(link.startsWith('https://teams.microsoft.us/l/message/')).toBe(true);
  });
});

describe('extractMessageTimestamp', () => {
  it('extracts from MessageId field', () => {
    expect(extractMessageTimestamp(sourceWithMessageId)).toBe('1705760000000');
  });

  it('extracts from ClientConversationId suffix', () => {
    expect(extractMessageTimestamp(sourceWithConvIdMessageId)).toBe('1705770000000');
  });

  it('falls back to parsing ISO timestamp', () => {
    const timestamp = extractMessageTimestamp(undefined, '2026-01-20T12:00:00.000Z');
    expect(timestamp).toBe(String(new Date('2026-01-20T12:00:00.000Z').getTime()));
  });

  it('returns undefined for missing data', () => {
    expect(extractMessageTimestamp(undefined)).toBeUndefined();
    expect(extractMessageTimestamp({})).toBeUndefined();
  });

  it('ignores invalid timestamp formats', () => {
    expect(extractMessageTimestamp(undefined, 'not-a-date')).toBeUndefined();
  });
});

describe('decodeBase64Guid', () => {
  it('decodes base64-encoded GUID correctly', () => {
    // '93qkaTtFGWpUHjyRafgdhg==' is a real base64-encoded GUID
    const result = decodeBase64Guid('93qkaTtFGWpUHjyRafgdhg==');
    expect(result).toBe('69a47af7-453b-6a19-541e-3c9169f81d86');
  });

  it('returns null for invalid base64', () => {
    expect(decodeBase64Guid('not-valid-base64!')).toBeNull();
  });

  it('returns null for wrong length', () => {
    // Too short (only 8 bytes when decoded)
    expect(decodeBase64Guid('AAAAAAAAAAA=')).toBeNull();
    // Too long (24 bytes when decoded)
    expect(decodeBase64Guid('AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==')).toBeNull();
  });

  it('returns lowercase GUID', () => {
    const result = decodeBase64Guid('93qkaTtFGWpUHjyRafgdhg==');
    expect(result).toBe(result?.toLowerCase());
  });
});

describe('parsePersonSuggestion', () => {
  it('parses complete person data', () => {
    const result = parsePersonSuggestion(personSuggestion);
    
    expect(result).not.toBeNull();
    expect(result!.id).toBe('a1b2c3d4-e5f6-7890-abcd-ef1234567890');
    expect(result!.mri).toBe('8:orgid:a1b2c3d4-e5f6-7890-abcd-ef1234567890');
    expect(result!.displayName).toBe('Smith, John');
    expect(result!.givenName).toBe('John');
    expect(result!.surname).toBe('Smith');
    expect(result!.email).toBe('john.smith@company.com');
    expect(result!.department).toBe('Engineering');
    expect(result!.jobTitle).toBe('Senior Engineer');
    expect(result!.companyName).toBe('Acme Corp');
  });

  it('handles minimal person data with GUID format', () => {
    const result = parsePersonSuggestion(personMinimal);
    
    expect(result).not.toBeNull();
    expect(result!.id).toBe('b1c2d3e4-f5a6-7890-bcde-1234567890ab');
    expect(result!.mri).toBe('8:orgid:b1c2d3e4-f5a6-7890-bcde-1234567890ab');
    expect(result!.displayName).toBe('Jane Doe');
    expect(result!.email).toBeUndefined();
  });

  it('decodes base64-encoded IDs', () => {
    const result = parsePersonSuggestion(personWithBase64Id);
    
    expect(result).not.toBeNull();
    expect(result!.id).toBe('69a47af7-453b-6a19-541e-3c9169f81d86');
    expect(result!.mri).toBe('8:orgid:69a47af7-453b-6a19-541e-3c9169f81d86');
    expect(result!.displayName).toBe('Rob MacDonald');
    expect(result!.email).toBe('rob@company.com');
  });

  it('extracts ID from tenant-qualified GUID format', () => {
    const result = parsePersonSuggestion({
      Id: 'a1b2c3d4-e5f6-7890-abcd-ef1234567890@tenant.onmicrosoft.com',
      DisplayName: 'Test User',
    });
    
    expect(result!.id).toBe('a1b2c3d4-e5f6-7890-abcd-ef1234567890');
  });

  it('returns null for missing ID', () => {
    expect(parsePersonSuggestion({ DisplayName: 'No ID' })).toBeNull();
  });

  it('returns null for invalid ID format', () => {
    expect(parsePersonSuggestion({ Id: 'invalid-format', DisplayName: 'Test' })).toBeNull();
  });
});

describe('parseV2Result', () => {
  it('parses complete search result', () => {
    const result = parseV2Result(searchResultItem);
    
    expect(result).not.toBeNull();
    expect(result!.type).toBe('message');
    expect(result!.content).toBe('Let me check the budget report for Q3');
    expect(result!.timestamp).toBe('2026-01-20T14:30:00.000Z');
    expect(result!.channelName).toBe('General');
    expect(result!.teamName).toBe('Finance Team');
    expect(result!.conversationId).toBe('19:abcdef123456@thread.tacv2');
    expect(result!.messageLink).toContain('teams.microsoft.com/l/message');
  });

  it('strips HTML from content', () => {
    const result = parseV2Result(searchResultWithHtml);
    
    expect(result).not.toBeNull();
    expect(result!.content).toBe("Meeting notes from & yesterday's call Action items:");
    expect(result!.content).not.toContain('<');
    expect(result!.content).not.toContain('>');
  });

  it('handles minimal result', () => {
    const result = parseV2Result(searchResultMinimal);
    
    expect(result).not.toBeNull();
    expect(result!.id).toBe('minimal-id');
    expect(result!.content).toBe('A short message here');
    expect(result!.conversationId).toBeUndefined();
    expect(result!.messageLink).toBeUndefined();
  });

  it('returns null for content too short', () => {
    expect(parseV2Result(searchResultTooShort)).toBeNull();
  });

  it('extracts conversationId from Extensions', () => {
    const result = parseV2Result(searchResultItem);
    expect(result!.conversationId).toBe('19:abcdef123456@thread.tacv2');
  });

  it('falls back to ClientThreadId for conversationId', () => {
    const result = parseV2Result(searchResultWithHtml);
    expect(result!.conversationId).toBe('19:meeting123@thread.v2');
  });

  it('generates messageLink with parentMessageId for thread replies', () => {
    const result = parseV2Result(searchResultThreadReply);
    
    expect(result).not.toBeNull();
    // Parent message ID from ClientConversationId;messageid=xxx
    expect(result!.messageLink).toContain('parentMessageId=1768919400000');
    // The message's own timestamp (from DateTimeReceived 2026-01-20T15:00:00.000Z)
    expect(result!.messageLink).toContain('/1768921200000');
  });

  it('generates messageLink with parentMessageId for top-level channel posts', () => {
    const result = parseV2Result(searchResultItem);
    
    expect(result).not.toBeNull();
    // Top-level post: parentMessageId equals the message's own timestamp
    // Per MS docs, parentMessageId is always required for channel links
    expect(result!.messageLink).toContain('parentMessageId=1768919400000');
    expect(result!.messageLink).toContain('createdTime=1768919400000');
  });

  it('generates messageLink with context for meeting chats', () => {
    const result = parseV2Result(searchResultWithHtml);
    
    expect(result).not.toBeNull();
    // Meeting chats need context parameter
    expect(result!.messageLink).toContain('context=');
  });
});

describe('parseJwtProfile', () => {
  it('parses complete JWT payload', () => {
    const profile = parseJwtProfile(jwtPayloadFull);
    
    expect(profile).not.toBeNull();
    expect(profile!.id).toBe('user-object-id-guid');
    expect(profile!.mri).toBe('8:orgid:user-object-id-guid');
    expect(profile!.email).toBe('rob.macdonald@company.com');
    expect(profile!.displayName).toBe('Macdonald, Rob');
    expect(profile!.givenName).toBe('Rob');
    expect(profile!.surname).toBe('Macdonald');
    expect(profile!.tenantId).toBe('tenant-id-guid');
  });

  it('handles minimal JWT payload', () => {
    const profile = parseJwtProfile(jwtPayloadMinimal);
    
    expect(profile).not.toBeNull();
    expect(profile!.id).toBe('another-user-guid');
    expect(profile!.displayName).toBe('Alice Smith');
    expect(profile!.email).toBe('');
    // Should parse from "Alice Smith" format
    expect(profile!.givenName).toBe('Alice');
    expect(profile!.surname).toBe('Smith');
  });

  it('parses "Surname, GivenName" format', () => {
    const profile = parseJwtProfile(jwtPayloadCommaName);
    
    expect(profile!.surname).toBe('Jones');
    expect(profile!.givenName).toBe('David');
  });

  it('parses "GivenName Surname" format', () => {
    const profile = parseJwtProfile(jwtPayloadSpaceName);
    
    expect(profile!.givenName).toBe('Sarah');
    expect(profile!.surname).toBe('Connor');
  });

  it('returns null for missing required fields', () => {
    expect(parseJwtProfile({})).toBeNull();
    expect(parseJwtProfile({ oid: 'id-only' })).toBeNull();
    expect(parseJwtProfile({ name: 'name-only' })).toBeNull();
  });

  it('prefers upn over other email fields', () => {
    const profile = parseJwtProfile(jwtPayloadFull);
    expect(profile!.email).toBe('rob.macdonald@company.com');
  });
});

describe('calculateTokenStatus', () => {
  const now = 1705846400000; // Fixed "now" for testing

  it('returns valid for unexpired token', () => {
    const expiry = now + 3600000; // 1 hour from now
    const status = calculateTokenStatus(expiry, now);
    
    expect(status.isValid).toBe(true);
    expect(status.minutesRemaining).toBe(60);
  });

  it('returns invalid for expired token', () => {
    const expiry = now - 60000; // 1 minute ago
    const status = calculateTokenStatus(expiry, now);
    
    expect(status.isValid).toBe(false);
    expect(status.minutesRemaining).toBe(0);
  });

  it('returns correct ISO date string', () => {
    const expiry = now + 3600000;
    const status = calculateTokenStatus(expiry, now);
    
    expect(status.expiresAt).toBe(new Date(expiry).toISOString());
  });

  it('rounds minutes correctly', () => {
    const status = calculateTokenStatus(now + 90000, now); // 1.5 minutes
    expect(status.minutesRemaining).toBe(2); // Rounds up
  });
});

describe('parseSearchResults', () => {
  it('parses EntitySets structure', () => {
    const { results, total } = parseSearchResults(
      searchEntitySetsResponse.EntitySets
    );
    
    expect(results).toHaveLength(2);
    expect(total).toBe(4307);
  });

  it('returns empty for undefined input', () => {
    const { results, total } = parseSearchResults(undefined);
    
    expect(results).toHaveLength(0);
    expect(total).toBeUndefined();
  });

  it('returns empty for non-array input', () => {
    const { results } = parseSearchResults(
      'not an array' as unknown as unknown[]
    );
    
    expect(results).toHaveLength(0);
  });

  it('filters out results with short content', () => {
    const entitySets = [{
      ResultSets: [{
        Results: [
          { Id: '1', HitHighlightedSummary: 'Valid content here' },
          { Id: '2', HitHighlightedSummary: 'Hi' }, // Too short
        ],
      }],
    }];
    
    const { results } = parseSearchResults(entitySets);
    expect(results).toHaveLength(1);
  });
});

describe('parsePeopleResults', () => {
  it('parses Groups/Suggestions structure', () => {
    const results = parsePeopleResults(peopleGroupsResponse.Groups);
    
    expect(results).toHaveLength(2);
    expect(results[0].displayName).toBe('Smith, John');
    expect(results[1].displayName).toBe('Jane Doe');
  });

  it('returns empty for undefined input', () => {
    expect(parsePeopleResults(undefined)).toHaveLength(0);
  });

  it('returns empty for non-array input', () => {
    expect(parsePeopleResults('not an array' as unknown as unknown[])).toHaveLength(0);
  });

  it('handles groups with no suggestions', () => {
    const groups = [{ Suggestions: [] }, { OtherField: 'value' }];
    expect(parsePeopleResults(groups)).toHaveLength(0);
  });
});

describe('extractObjectId', () => {
  it('extracts GUID from MRI format', () => {
    expect(extractObjectId('8:orgid:ab76f827-27e2-4c67-a765-f1a53145fa24'))
      .toBe('ab76f827-27e2-4c67-a765-f1a53145fa24');
  });

  it('extracts GUID from Skype ID format (without 8: prefix)', () => {
    expect(extractObjectId('orgid:ab76f827-27e2-4c67-a765-f1a53145fa24'))
      .toBe('ab76f827-27e2-4c67-a765-f1a53145fa24');
  });

  it('extracts GUID from ID with tenant format', () => {
    expect(extractObjectId('5817f485-f870-46eb-bbc4-de216babac62@56b731a8-a2ac-4c32-bf6b-616810e913c6'))
      .toBe('5817f485-f870-46eb-bbc4-de216babac62');
  });

  it('returns raw GUID unchanged', () => {
    expect(extractObjectId('ab76f827-27e2-4c67-a765-f1a53145fa24'))
      .toBe('ab76f827-27e2-4c67-a765-f1a53145fa24');
  });

  it('normalises to lowercase', () => {
    expect(extractObjectId('AB76F827-27E2-4C67-A765-F1A53145FA24'))
      .toBe('ab76f827-27e2-4c67-a765-f1a53145fa24');
  });

  it('decodes base64-encoded GUID', () => {
    expect(extractObjectId('93qkaTtFGWpUHjyRafgdhg=='))
      .toBe('69a47af7-453b-6a19-541e-3c9169f81d86');
  });

  it('decodes base64 GUID from MRI format', () => {
    expect(extractObjectId('8:orgid:93qkaTtFGWpUHjyRafgdhg=='))
      .toBe('69a47af7-453b-6a19-541e-3c9169f81d86');
  });

  it('decodes base64 GUID from Skype ID format', () => {
    expect(extractObjectId('orgid:93qkaTtFGWpUHjyRafgdhg=='))
      .toBe('69a47af7-453b-6a19-541e-3c9169f81d86');
  });

  it('returns null for invalid formats', () => {
    expect(extractObjectId('')).toBeNull();
    expect(extractObjectId('not-a-guid')).toBeNull();
    expect(extractObjectId('8:orgid:invalid')).toBeNull();
    expect(extractObjectId('orgid:invalid')).toBeNull();
    expect(extractObjectId('missing-sections-1234')).toBeNull();
  });
});

describe('buildOneOnOneConversationId', () => {
  const userId1 = 'ab76f827-27e2-4c67-a765-f1a53145fa24';
  const userId2 = '5817f485-f870-46eb-bbc4-de216babac62';

  it('builds conversation ID with sorted user IDs', () => {
    // userId2 ('5817...') comes before userId1 ('ab76...') alphabetically
    const result = buildOneOnOneConversationId(userId1, userId2);
    expect(result).toBe('19:5817f485-f870-46eb-bbc4-de216babac62_ab76f827-27e2-4c67-a765-f1a53145fa24@unq.gbl.spaces');
  });

  it('produces same result regardless of argument order', () => {
    const result1 = buildOneOnOneConversationId(userId1, userId2);
    const result2 = buildOneOnOneConversationId(userId2, userId1);
    expect(result1).toBe(result2);
  });

  it('handles MRI format input', () => {
    const mri1 = `8:orgid:${userId1}`;
    const mri2 = `8:orgid:${userId2}`;
    const result = buildOneOnOneConversationId(mri1, mri2);
    expect(result).toBe('19:5817f485-f870-46eb-bbc4-de216babac62_ab76f827-27e2-4c67-a765-f1a53145fa24@unq.gbl.spaces');
  });

  it('handles ID with tenant format', () => {
    const idWithTenant = '5817f485-f870-46eb-bbc4-de216babac62@56b731a8-a2ac-4c32-bf6b-616810e913c6';
    const result = buildOneOnOneConversationId(userId1, idWithTenant);
    expect(result).toBe('19:5817f485-f870-46eb-bbc4-de216babac62_ab76f827-27e2-4c67-a765-f1a53145fa24@unq.gbl.spaces');
  });

  it('handles base64-encoded GUID input', () => {
    // '93qkaTtFGWpUHjyRafgdhg==' decodes to '69a47af7-453b-6a19-541e-3c9169f81d86'
    const base64Id = '93qkaTtFGWpUHjyRafgdhg==';
    const result = buildOneOnOneConversationId(base64Id, userId2);
    // '5817...' < '69a4...' so 5817 comes first
    expect(result).toBe('19:5817f485-f870-46eb-bbc4-de216babac62_69a47af7-453b-6a19-541e-3c9169f81d86@unq.gbl.spaces');
  });

  it('returns null for invalid input', () => {
    expect(buildOneOnOneConversationId('invalid', userId2)).toBeNull();
    expect(buildOneOnOneConversationId(userId1, 'invalid')).toBeNull();
    expect(buildOneOnOneConversationId('', '')).toBeNull();
  });
});

describe('extractActivityTimestamp', () => {
  it('prefers originalarrivaltime when present', () => {
    const msg = {
      originalarrivaltime: '2024-01-15T10:30:00.000Z',
      composetime: '2024-01-15T10:29:00.000Z',
      id: '1705315800000',
    };
    expect(extractActivityTimestamp(msg)).toBe('2024-01-15T10:30:00.000Z');
  });

  it('falls back to composetime when originalarrivaltime is missing', () => {
    const msg = {
      composetime: '2024-01-15T10:29:00.000Z',
      id: '1705315800000',
    };
    expect(extractActivityTimestamp(msg)).toBe('2024-01-15T10:29:00.000Z');
  });

  it('parses numeric id as timestamp when no time fields present', () => {
    const msg = {
      id: '1705315800000', // 2024-01-15T10:30:00.000Z
    };
    const result = extractActivityTimestamp(msg);
    expect(result).toBe(new Date(1705315800000).toISOString());
  });

  it('returns null for non-numeric id when no time fields present', () => {
    const msg = {
      id: 'abc-not-a-number',
    };
    expect(extractActivityTimestamp(msg)).toBeNull();
  });

  it('returns null for empty message object', () => {
    expect(extractActivityTimestamp({})).toBeNull();
  });

  it('returns null when id is undefined', () => {
    const msg = {
      originalarrivaltime: undefined,
      composetime: undefined,
    };
    expect(extractActivityTimestamp(msg)).toBeNull();
  });

  it('handles zero id correctly (returns null)', () => {
    const msg = {
      id: '0',
    };
    // Zero is not a valid timestamp
    expect(extractActivityTimestamp(msg)).toBeNull();
  });

  it('handles negative id correctly (returns null)', () => {
    const msg = {
      id: '-1705315800000',
    };
    // Negative timestamps are invalid
    expect(extractActivityTimestamp(msg)).toBeNull();
  });
});

describe('markdownToTeamsHtml', () => {
  it('wraps plain text in paragraph tags', () => {
    expect(markdownToTeamsHtml('Hello world')).toBe('<p>Hello world</p>');
  });

  it('escapes HTML special characters', () => {
    expect(markdownToTeamsHtml('1 < 2 & 3 > 0')).toBe('<p>1 &lt; 2 &amp; 3 &gt; 0</p>');
  });

  it('converts bold markdown', () => {
    expect(markdownToTeamsHtml('This is **bold** text')).toBe('<p>This is <b>bold</b> text</p>');
    expect(markdownToTeamsHtml('This is __bold__ text')).toBe('<p>This is <b>bold</b> text</p>');
  });

  it('converts italic markdown', () => {
    expect(markdownToTeamsHtml('This is *italic* text')).toBe('<p>This is <i>italic</i> text</p>');
  });

  it('converts strikethrough markdown', () => {
    expect(markdownToTeamsHtml('This is ~~deleted~~ text')).toBe('<p>This is <s>deleted</s> text</p>');
  });

  it('converts inline code', () => {
    expect(markdownToTeamsHtml('Use `console.log()` here')).toBe('<p>Use <code>console.log()</code> here</p>');
  });

  it('does not process markdown inside inline code', () => {
    expect(markdownToTeamsHtml('Use `**not bold**` here')).toBe('<p>Use <code>**not bold**</code> here</p>');
  });

  it('escapes HTML inside inline code', () => {
    expect(markdownToTeamsHtml('Use `<div>` tag')).toBe('<p>Use <code>&lt;div&gt;</code> tag</p>');
  });

  it('converts fenced code blocks', () => {
    expect(markdownToTeamsHtml('```\nconst x = 1;\n```')).toBe('<pre><code>const x = 1;</code></pre>');
  });

  it('converts fenced code blocks with language', () => {
    expect(markdownToTeamsHtml('```js\nconst x = 1;\n```')).toBe('<pre><code>const x = 1;</code></pre>');
  });

  it('escapes HTML inside code blocks', () => {
    expect(markdownToTeamsHtml('```\n<div>test</div>\n```')).toBe('<pre><code>&lt;div&gt;test&lt;/div&gt;</code></pre>');
  });

  it('handles text before and after code blocks', () => {
    const input = 'Before\n\n```\ncode\n```\n\nAfter';
    expect(markdownToTeamsHtml(input)).toBe('<p>Before</p><pre><code>code</code></pre><p>After</p>');
  });

  it('converts single newlines to br tags', () => {
    expect(markdownToTeamsHtml('Line 1\nLine 2')).toBe('<p>Line 1<br>Line 2</p>');
  });

  it('converts double newlines to separate paragraphs', () => {
    expect(markdownToTeamsHtml('Para 1\n\nPara 2')).toBe('<p>Para 1</p><p>Para 2</p>');
  });

  it('converts unordered lists', () => {
    expect(markdownToTeamsHtml('- Item 1\n- Item 2\n- Item 3')).toBe(
      '<ul><li>Item 1</li><li>Item 2</li><li>Item 3</li></ul>'
    );
  });

  it('converts unordered lists with * marker', () => {
    expect(markdownToTeamsHtml('* Item 1\n* Item 2')).toBe(
      '<ul><li>Item 1</li><li>Item 2</li></ul>'
    );
  });

  it('converts ordered lists', () => {
    expect(markdownToTeamsHtml('1. First\n2. Second\n3. Third')).toBe(
      '<ol><li>First</li><li>Second</li><li>Third</li></ol>'
    );
  });

  it('handles inline formatting inside list items', () => {
    expect(markdownToTeamsHtml('- **Bold** item\n- *Italic* item')).toBe(
      '<ul><li><b>Bold</b> item</li><li><i>Italic</i> item</li></ul>'
    );
  });

  it('handles combined formatting', () => {
    const input = '**Bold** and *italic* and `code`';
    expect(markdownToTeamsHtml(input)).toBe('<p><b>Bold</b> and <i>italic</i> and <code>code</code></p>');
  });

  it('handles complex multi-paragraph message', () => {
    const input = 'Hello **team**!\n\nHere are the updates:\n\n- Item 1\n- Item 2\n\nThanks!';
    expect(markdownToTeamsHtml(input)).toBe(
      '<p>Hello <b>team</b>!</p><p>Here are the updates:</p><ul><li>Item 1</li><li>Item 2</li></ul><p>Thanks!</p>'
    );
  });

  it('returns empty paragraph for empty string', () => {
    expect(markdownToTeamsHtml('')).toBe('<p></p>');
  });

  it('handles whitespace-only input', () => {
    expect(markdownToTeamsHtml('   ')).toBe('<p></p>');
  });
});

describe('formatTranscriptText', () => {
  it('formats entries with speaker names and timestamps', () => {
    const entries = [
      { startTime: '00:00:01.000', endTime: '00:00:05.000', speaker: 'Alice', text: 'Hello everyone.' },
      { startTime: '00:00:06.000', endTime: '00:00:10.000', speaker: 'Bob', text: 'Hi Alice!' },
    ];
    const result = formatTranscriptText(entries);
    expect(result).toBe(
      '[00:00:01.000] Alice:\nHello everyone.\n\n[00:00:06.000] Bob:\nHi Alice!'
    );
  });

  it('merges consecutive entries from the same speaker', () => {
    const entries = [
      { startTime: '00:00:01.000', endTime: '00:00:03.000', speaker: 'Alice', text: 'First part.' },
      { startTime: '00:00:03.000', endTime: '00:00:06.000', speaker: 'Alice', text: 'Second part.' },
      { startTime: '00:00:07.000', endTime: '00:00:10.000', speaker: 'Bob', text: 'Response.' },
    ];
    const result = formatTranscriptText(entries);
    expect(result).toBe(
      '[00:00:01.000] Alice:\nFirst part. Second part.\n\n[00:00:07.000] Bob:\nResponse.'
    );
  });

  it('handles entries without speaker names', () => {
    const entries = [
      { startTime: '00:00:01.000', endTime: '00:00:05.000', speaker: '', text: 'Unknown speaker.' },
    ];
    const result = formatTranscriptText(entries);
    expect(result).toBe('[00:00:01.000]\nUnknown speaker.');
  });

  it('returns empty string for empty entries', () => {
    expect(formatTranscriptText([])).toBe('');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Channel Parsing
// ─────────────────────────────────────────────────────────────────────────────

describe('parseChannelSuggestion', () => {
  it('parses a complete channel suggestion', () => {
    const result = parseChannelSuggestion({
      Name: 'General',
      ThreadId: '19:abc@thread.tacv2',
      TeamName: 'Engineering',
      GroupId: 'group-guid-123',
      ChannelType: 'Standard',
      Description: 'Main channel',
    });
    expect(result).toEqual({
      channelId: '19:abc@thread.tacv2',
      channelName: 'General',
      teamName: 'Engineering',
      teamId: 'group-guid-123',
      channelType: 'Standard',
      description: 'Main channel',
    });
  });

  it('returns null when required fields are missing', () => {
    expect(parseChannelSuggestion({ Name: 'General' })).toBeNull();
    expect(parseChannelSuggestion({ Name: 'General', ThreadId: '19:abc@thread.tacv2' })).toBeNull();
    expect(parseChannelSuggestion({})).toBeNull();
  });

  it('defaults channelType to Standard when missing', () => {
    const result = parseChannelSuggestion({
      Name: 'General',
      ThreadId: '19:abc@thread.tacv2',
      TeamName: 'Engineering',
      GroupId: 'group-guid-123',
    });
    expect(result?.channelType).toBe('Standard');
  });
});

describe('parseChannelResults', () => {
  it('parses channel suggestions from groups structure', () => {
    const groups = [
      {
        Suggestions: [
          {
            EntityType: 'ChannelSuggestion',
            Name: 'General',
            ThreadId: '19:abc@thread.tacv2',
            TeamName: 'Engineering',
            GroupId: 'group-guid-123',
          },
          {
            EntityType: 'PersonSuggestion',
            DisplayName: 'John',
          },
        ],
      },
    ];
    const results = parseChannelResults(groups);
    expect(results).toHaveLength(1);
    expect(results[0].channelName).toBe('General');
  });

  it('returns empty array for undefined/non-array input', () => {
    expect(parseChannelResults(undefined)).toEqual([]);
    expect(parseChannelResults([])).toEqual([]);
  });

  it('skips non-ChannelSuggestion entities', () => {
    const groups = [
      {
        Suggestions: [
          { EntityType: 'PersonSuggestion', DisplayName: 'John' },
        ],
      },
    ];
    expect(parseChannelResults(groups)).toEqual([]);
  });
});

describe('parseTeamsList', () => {
  it('parses teams with channels', () => {
    const data = {
      teams: [
        {
          id: '19:team1@thread.tacv2',
          displayName: 'Engineering',
          description: 'Eng team',
          channels: [
            {
              id: '19:channel1@thread.tacv2',
              displayName: 'General',
              groupId: 'group-1',
              channelType: 0,
            },
            {
              id: '19:channel2@thread.tacv2',
              displayName: 'Private Channel',
              groupId: 'group-1',
              channelType: 1,
            },
          ],
        },
      ],
    };
    const results = parseTeamsList(data);
    expect(results).toHaveLength(1);
    expect(results[0].teamName).toBe('Engineering');
    expect(results[0].channels).toHaveLength(2);
    expect(results[0].channels[0].channelType).toBe('Standard');
    expect(results[0].channels[1].channelType).toBe('Private');
    expect(results[0].channels[0].isMember).toBe(true);
  });

  it('maps channelType 2 to Shared', () => {
    const data = {
      teams: [
        {
          id: '19:team1@thread.tacv2',
          displayName: 'Eng',
          channels: [
            { id: '19:ch@thread.tacv2', displayName: 'Shared', groupId: 'g', channelType: 2 },
          ],
        },
      ],
    };
    expect(parseTeamsList(data)[0].channels[0].channelType).toBe('Shared');
  });

  it('returns empty array for undefined/missing data', () => {
    expect(parseTeamsList(undefined)).toEqual([]);
    expect(parseTeamsList({})).toEqual([]);
    expect(parseTeamsList({ teams: 'not-array' } as unknown as Record<string, unknown>)).toEqual([]);
  });

  it('skips teams without required fields', () => {
    const data = {
      teams: [
        { id: '19:team@thread.tacv2' }, // missing displayName
        { displayName: 'No ID' },       // missing id
      ],
    };
    expect(parseTeamsList(data)).toEqual([]);
  });
});

describe('filterChannelsByName', () => {
  const teams = [
    {
      teamId: '19:team1@thread.tacv2',
      teamName: 'Engineering',
      threadId: '19:team1@thread.tacv2',
      channels: [
        { channelId: '19:ch1@thread.tacv2', channelName: 'General', teamName: 'Engineering', teamId: 'g1', channelType: 'Standard' },
        { channelId: '19:ch2@thread.tacv2', channelName: 'Design Review', teamName: 'Engineering', teamId: 'g1', channelType: 'Standard' },
      ],
    },
    {
      teamId: '19:team2@thread.tacv2',
      teamName: 'Marketing',
      threadId: '19:team2@thread.tacv2',
      channels: [
        { channelId: '19:ch3@thread.tacv2', channelName: 'General', teamName: 'Marketing', teamId: 'g2', channelType: 'Standard' },
      ],
    },
  ];

  it('filters channels by partial name match (case-insensitive)', () => {
    const results = filterChannelsByName(teams, 'general');
    expect(results).toHaveLength(2);
  });

  it('returns matching channels across teams', () => {
    const results = filterChannelsByName(teams, 'design');
    expect(results).toHaveLength(1);
    expect(results[0].channelName).toBe('Design Review');
  });

  it('returns empty array when no match', () => {
    expect(filterChannelsByName(teams, 'nonexistent')).toEqual([]);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Virtual Conversation Parsing
// ─────────────────────────────────────────────────────────────────────────────

describe('parseVirtualConversationMessage', () => {
  const savedPattern = /_M_(\d+)$/;
  const followedPattern = /_P_(\d+)_Threads$/;

  it('parses a saved message with all fields', () => {
    const msg = {
      id: '1705760000000',
      messagetype: 'RichText/Html',
      content: '<p>Hello world</p>',
      from: '8:orgid:user-guid',
      imdisplayname: 'John Smith',
      originalarrivaltime: '2026-01-20T12:00:00.000Z',
      clumpId: '19:abc@thread.tacv2',
      secondaryReferenceId: 'T_19:abc@thread.tacv2_M_1705760000000',
    };
    const result = parseVirtualConversationMessage(msg, savedPattern);
    expect(result).not.toBeNull();
    expect(result!.id).toBe('1705760000000');
    expect(result!.content).toBe('Hello world');
    expect(result!.sourceConversationId).toBe('19:abc@thread.tacv2');
    expect(result!.sourceReferenceId).toBe('1705760000000');
    expect(result!.sender.displayName).toBe('John Smith');
  });

  it('parses a followed thread reference', () => {
    const msg = {
      id: '1705770000000',
      messagetype: 'Text',
      content: 'Thread content',
      from: '8:orgid:user-guid',
      originalarrivaltime: '2026-01-20T13:00:00.000Z',
      clumpId: '19:xyz@thread.tacv2',
      secondaryReferenceId: 'T_19:xyz@thread.tacv2_P_1705770000000_Threads',
    };
    const result = parseVirtualConversationMessage(msg, followedPattern);
    expect(result).not.toBeNull();
    expect(result!.sourceReferenceId).toBe('1705770000000');
  });

  it('returns null for Control messages', () => {
    const msg = {
      id: '123',
      messagetype: 'Control/Typing',
      originalarrivaltime: '2026-01-20T12:00:00.000Z',
    };
    expect(parseVirtualConversationMessage(msg, savedPattern)).toBeNull();
  });

  it('returns null for messages without id', () => {
    const msg = {
      messagetype: 'Text',
      originalarrivaltime: '2026-01-20T12:00:00.000Z',
    };
    expect(parseVirtualConversationMessage(msg, savedPattern)).toBeNull();
  });

  it('returns null for messages without timestamp and non-numeric id', () => {
    const msg = {
      id: 'not-a-number',
      messagetype: 'Text',
    };
    expect(parseVirtualConversationMessage(msg, savedPattern)).toBeNull();
  });

  it('handles missing secondaryReferenceId gracefully', () => {
    const msg = {
      id: '123',
      messagetype: 'Text',
      content: 'Hello',
      from: '8:orgid:user',
      originalarrivaltime: '2026-01-20T12:00:00.000Z',
      clumpId: '19:abc@thread.tacv2',
    };
    const result = parseVirtualConversationMessage(msg, savedPattern);
    expect(result).not.toBeNull();
    expect(result!.sourceReferenceId).toBeUndefined();
    expect(result!.messageLink).toBeUndefined();
  });

  it('builds message link when linkContext is provided', () => {
    const msg = {
      id: '123',
      messagetype: 'Text',
      content: 'Hello',
      from: '8:orgid:user',
      originalarrivaltime: '2026-01-20T12:00:00.000Z',
      clumpId: '19:abc@thread.tacv2',
      secondaryReferenceId: 'T_19:abc@thread.tacv2_M_1705760000000',
    };
    const result = parseVirtualConversationMessage(msg, savedPattern, {
      tenantId: 'tenant-123',
      teamsBaseUrl: 'https://teams.microsoft.com',
    });
    expect(result).not.toBeNull();
    expect(result!.messageLink).toBeDefined();
    expect(result!.messageLink).toContain('teams.microsoft.com');
  });

  it('extracts links from HTML content', () => {
    const msg = {
      id: '123',
      messagetype: 'RichText/Html',
      content: '<a href="https://example.com">Link</a>',
      from: '8:orgid:user',
      originalarrivaltime: '2026-01-20T12:00:00.000Z',
      clumpId: '19:abc@thread.tacv2',
    };
    const result = parseVirtualConversationMessage(msg, savedPattern);
    expect(result).not.toBeNull();
    expect(result!.links).toHaveLength(1);
    expect(result!.links![0].url).toBe('https://example.com');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// hasMarkdownFormatting
// ─────────────────────────────────────────────────────────────────────────────

describe('hasMarkdownFormatting', () => {
  it('detects bold formatting', () => {
    expect(hasMarkdownFormatting('**bold**')).toBe(true);
    expect(hasMarkdownFormatting('__bold__')).toBe(true);
  });

  it('detects italic formatting', () => {
    expect(hasMarkdownFormatting('*italic*')).toBe(true);
  });

  it('detects strikethrough', () => {
    expect(hasMarkdownFormatting('~~strike~~')).toBe(true);
  });

  it('detects inline code', () => {
    expect(hasMarkdownFormatting('use `code` here')).toBe(true);
  });

  it('detects code blocks', () => {
    expect(hasMarkdownFormatting('```\ncode\n```')).toBe(true);
  });

  it('detects unordered lists', () => {
    expect(hasMarkdownFormatting('- item one')).toBe(true);
    expect(hasMarkdownFormatting('* item one')).toBe(true);
  });

  it('detects ordered lists', () => {
    expect(hasMarkdownFormatting('1. first item')).toBe(true);
    expect(hasMarkdownFormatting('2) second item')).toBe(true);
  });

  it('detects newlines as needing conversion', () => {
    expect(hasMarkdownFormatting('line one\nline two')).toBe(true);
  });

  it('returns false for plain text', () => {
    expect(hasMarkdownFormatting('just plain text')).toBe(false);
    expect(hasMarkdownFormatting('hello world')).toBe(false);
    expect(hasMarkdownFormatting('')).toBe(false);
  });
});
