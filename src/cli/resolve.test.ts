/**
 * CLI resolution tests.
 *
 * These run against the real tool definitions, so they also guard the
 * assumption the CLI rests on: that commands, flags and positional arguments
 * can be derived from the schemas.
 */

import { describe, it, expect, vi } from 'vitest';
import { fileURLToPath } from 'node:url';
import { dirname } from 'node:path';
import { isatty } from 'node:tty';
import { getToolDefinitions } from '../tools/registry.js';
import {
  UsageError,
  parseToolArgs,
  positionalProp,
  resolveToolName,
  shortName,
} from './resolve.js';
import { firstSentence, renderToolHelp } from './help.js';

vi.mock('node:tty', () => ({ isatty: vi.fn(() => false) }));

const tools = getToolDefinitions();
const names = tools.map(tool => tool.name);
const tool = (name: string) => tools.find(t => t.name === name)!;

describe('resolveToolName', () => {
  it.each([
    ['teams_search', 'teams_search'],
    ['search', 'teams_search'],
    ['thread', 'teams_get_thread'],
    ['get_thread', 'teams_get_thread'],
    ['people', 'teams_search_people'],
    ['send', 'teams_send_message'],
    ['send_message', 'teams_send_message'],
    ['save', 'teams_save_message'],
    ['unsave', 'teams_unsave_message'],
    ['markread', 'teams_mark_read'],
    ['channel', 'teams_find_channel'],
    ['chat', 'teams_get_chat'],
    ['chats', 'teams_list_chats'],
    ['message', 'teams_get_message'],
    ['assignment', 'teams_get_assignment'],
    ['assignments', 'teams_list_assignments'],
    ['SEARCH', 'teams_search'],
  ])('resolves %s', (command, expected) => {
    expect(resolveToolName(names, command)).toBe(expected);
  });

  it('rejects an unknown command', () => {
    expect(() => resolveToolName(names, 'nope')).toThrow(UsageError);
  });

  it('rejects an ambiguous command and names the candidates', () => {
    expect(() => resolveToolName(names, 'reaction')).toThrow(/teams_add_reaction/);
  });
});

describe('shortName', () => {
  it('gives every tool a label that resolves back to it', () => {
    for (const name of names) {
      expect(resolveToolName(names, shortName(names, name))).toBe(name);
    }
  });
});

describe('positionalProp', () => {
  it('is the first required string', () => {
    expect(positionalProp(tool('teams_search'))).toBe('query');
    expect(positionalProp(tool('teams_send_message'))).toBe('content');
  });

  it('is absent when nothing is required', () => {
    expect(positionalProp(tool('teams_get_favorites'))).toBeNull();
  });
});

describe('parseToolArgs', () => {
  const parse = (name: string, argv: string[]) => parseToolArgs(tool(name), argv);

  it('maps a bare value to the positional parameter', () => {
    expect(parse('teams_search', ['hello']).args).toEqual({ query: 'hello' });
  });

  it('coerces numbers from the schema type and leaves ids as strings', () => {
    const { args } = parse('teams_get_thread', ['--conversationId', '19:abc', '--limit', '25']);
    expect(args).toEqual({ conversationId: '19:abc', limit: 25 });
  });

  it('accepts booleans as a bare flag and as --flag true|false', () => {
    expect(parse('teams_get_thread', ['--markRead']).args.markRead).toBe(true);
    expect(parse('teams_get_thread', ['--markRead', 'true']).args.markRead).toBe(true);
    expect(parse('teams_get_thread', ['--markRead', 'false']).args.markRead).toBe(false);
  });

  it('parses arrays as JSON or a comma-separated list', () => {
    expect(parse('teams_create_group_chat', ['--userIds', '["a","b"]']).args.userIds).toEqual(['a', 'b']);
    expect(parse('teams_create_group_chat', ['--userIds', 'a, b']).args.userIds).toEqual(['a', 'b']);
  });

  it('applies flag aliases', () => {
    const { args } = parse('teams_send_message', ['hi', '--to', '19:abc']);
    expect(args).toEqual({ content: 'hi', conversationId: '19:abc' });
  });

  it('prefers an explicit flag over the positional', () => {
    expect(parse('teams_search', ['bare', '--query', 'flag']).args.query).toBe('flag');
  });

  it('extracts the global flags without passing them to the tool', () => {
    const parsed = parse('teams_search', ['hi', '--json']);
    expect(parsed.json).toBe(true);
    expect(parsed.args).toEqual({ query: 'hi' });
    expect(parse('teams_search', ['--help']).help).toBe(true);
  });

  it.each([
    ['an unknown flag', 'teams_search', ['hi', '--nope', 'x']],
    ['a non-numeric number', 'teams_search', ['hi', '--size', 'abc']],
    ['a missing required parameter', 'teams_send_message', []],
    ['a second bare value', 'teams_search', ['one', 'two']],
    ['a value for a command that takes none', 'teams_get_favorites', ['oops']],
  ])('rejects %s', (_label, name, argv) => {
    expect(() => parse(name, argv)).toThrow(UsageError);
  });

  it('does not require parameters that --help only wants to read about', () => {
    expect(parse('teams_send_message', ['--help']).help).toBe(true);
  });
});

describe('parseToolArgs --content-file', () => {
  const sendMessage = tool('teams_send_message');
  const editMessage = tool('teams_edit_message');

  it('reads multi-line content for send_message and keeps later flags parsed', () => {
    const content = 'First line „quoted” →\n\nSecond paragraph with diacritics: ăâîșț';
    const seen: string[] = [];
    const { args } = parseToolArgs(
      sendMessage,
      ['--content-file', 'msg.md', '--files', 'a,b', '--conversationId', '19:xyz'],
      'send_message',
      (pathOrDash) => { seen.push(pathOrDash); return content; }
    );

    expect(seen).toEqual(['msg.md']);
    expect(args.content).toBe(content);
    expect(args.files).toEqual(['a', 'b']);
    expect(args.conversationId).toBe('19:xyz');
  });

  it('accepts --content-file for edit_message too', () => {
    const { args } = parseToolArgs(
      editMessage,
      ['--conversationId', '19:abc', '--messageId', 'msg-1', '--content-file', 'edit.md'],
      'edit_message',
      () => 'replacement body'
    );

    expect(args.content).toBe('replacement body');
  });

  it('normalises BOM, CRLF and trailing newlines but keeps internal blank lines', () => {
    const { args } = parseToolArgs(
      sendMessage,
      ['--content-file', 'msg.md'],
      'send_message',
      () => '\uFEFFline one\r\n\r\nline two  \r\n\r\n'
    );

    expect(args.content).toBe('line one\n\nline two  ');
  });

  it('passes "-" through to the reader for stdin', () => {
    const seen: string[] = [];
    parseToolArgs(sendMessage, ['--content-file', '-'], 'send_message', (pathOrDash) => {
      seen.push(pathOrDash);
      return 'from stdin';
    });

    expect(seen).toEqual(['-']);
  });

  it('rejects --content-file together with --content', () => {
    expect(() => parseToolArgs(
      sendMessage,
      ['--content', 'hi', '--content-file', 'msg.md'],
      'send_message',
      () => 'file body'
    )).toThrow(/not both/);
  });

  it('rejects --content-file together with a positional content value', () => {
    expect(() => parseToolArgs(
      sendMessage,
      ['positional body', '--content-file', 'msg.md'],
      'send_message',
      () => 'file body'
    )).toThrow(/not both/);
  });

  it('turns a read failure into a UsageError naming the path', () => {
    let error: unknown;
    try {
      parseToolArgs(
        sendMessage,
        ['--content-file', 'missing.md'],
        'send_message',
        () => { throw new Error('boom'); }
      );
    } catch (e) {
      error = e;
    }
    expect(error).toBeInstanceOf(UsageError);
    expect((error as Error).message).toBe("Cannot read --content-file 'missing.md': boom");
  });

  it('rejects an empty or whitespace-only file', () => {
    expect(() => parseToolArgs(
      sendMessage,
      ['--content-file', 'empty.md'],
      'send_message',
      () => '  \n\n\t\n'
    )).toThrow(UsageError);
  });

  it('is an unknown flag for tools without a content parameter', () => {
    expect(() => parseToolArgs(
      tool('teams_search'),
      ['hi', '--content-file', 'msg.md'],
      'search',
      () => 'body'
    )).toThrow(UsageError);
  });

  it('reads a real file with the default reader', () => {
    const fixture = fileURLToPath(new URL('../__fixtures__/multiline-message.md', import.meta.url));
    const { args } = parseToolArgs(sendMessage, ['--content-file', fixture]);

    expect(args.content).toContain('diacritice românești');
    expect(args.content).toContain('**Acesta este bold.**');
    expect(args.content).toContain('  - element imbricat');
  });

  it('decodes a UTF-16LE file (Windows PowerShell output) instead of sending NUL garbage', () => {
    const fixture = fileURLToPath(new URL('../__fixtures__/utf16-message.md', import.meta.url));
    const { args } = parseToolArgs(sendMessage, ['--content-file', fixture]);

    expect(args.content).toContain('diacritice: ă â î ș ț');
    expect(args.content).not.toContain('\u0000');
  });

  it('rejects a file that is not valid UTF-8', () => {
    const fixture = fileURLToPath(new URL('../__fixtures__/ansi-message.md', import.meta.url));
    expect(() => parseToolArgs(sendMessage, ['--content-file', fixture]))
      .toThrow(UsageError);
    expect(() => parseToolArgs(sendMessage, ['--content-file', fixture]))
      .toThrow(/UTF-8/);
  });

  it('rejects content containing NUL bytes (UTF-16 saved without a BOM)', () => {
    expect(() => parseToolArgs(
      sendMessage,
      ['--content-file', 'x.md'],
      'send_message',
      () => 'a\u0000b'
    )).toThrow(UsageError);
  });

  it('resolves relative --content-file paths against INIT_CWD, not the package root', () => {
    // `npm run cli` sets cwd to the package root but keeps the caller's
    // directory in INIT_CWD - a relative path must resolve there.
    const fixtureDir = dirname(fileURLToPath(new URL('../__fixtures__/multiline-message.md', import.meta.url)));
    vi.stubEnv('INIT_CWD', fixtureDir);
    try {
      const { args } = parseToolArgs(sendMessage, ['--content-file', 'multiline-message.md']);
      expect(args.content).toContain('diacritice românești');
    } finally {
      vi.unstubAllEnvs();
    }
  });

  it('refuses "-" on a TTY instead of blocking on stdin', () => {
    vi.mocked(isatty).mockReturnValueOnce(true);
    expect(() => parseToolArgs(sendMessage, ['--content-file', '-']))
      .toThrow(/reads stdin/);
  });

  it('does not read the file when --help is present', () => {
    const result = parseToolArgs(
      sendMessage,
      ['--content-file', '-', '--help'],
      'send_message',
      () => { throw new Error('must not read'); }
    );
    expect(result.help).toBe(true);
  });

  it('rejects the conflict before reading the file', () => {
    let calls = 0;
    expect(() => parseToolArgs(
      sendMessage,
      ['--content', 'hi', '--content-file', 'f.md'],
      'send_message',
      () => { calls++; return 'x'; }
    )).toThrow(/not both/);
    expect(calls).toBe(0);
  });

  it('rejects --content-file together with --content on edit_message too', () => {
    expect(() => parseToolArgs(
      editMessage,
      ['--conversationId', '19:abc', '--messageId', 'msg-1', '--content', 'x', '--content-file', 'f.md'],
      'edit_message',
      () => 'y'
    )).toThrow(/not both/);
  });

  it('mentions --content-file in usage errors for content tools only', () => {
    expect(() => parseToolArgs(sendMessage, ['--nope'], 'send_message'))
      .toThrow(/--content-file <path\|->/);
    try {
      parseToolArgs(tool('teams_search'), ['--nope']);
      expect.unreachable('expected a UsageError');
    } catch (e) {
      expect((e as Error).message).not.toContain('content-file');
    }
  });

  it('lists --content-file in per-command help for content tools only', () => {
    expect(renderToolHelp(tools, sendMessage)).toContain('--content-file');
    expect(renderToolHelp(tools, editMessage)).toContain('--content-file');
    expect(renderToolHelp(tools, tool('teams_search'))).not.toContain('content-file');
  });
});

describe('firstSentence', () => {
  it('stops at the end of the first sentence', () => {
    expect(firstSentence('Find Teams channels by name. Searches both (1) channels...'))
      .toBe('Find Teams channels by name');
  });

  it('does not treat an abbreviation as the end', () => {
    expect(firstSentence('Get profiles by MRI (e.g. "8:orgid:uuid"). Returns display name.'))
      .toBe('Get profiles by MRI (e.g. "8:orgid:uuid")');
  });

  it('gives every tool a non-empty summary', () => {
    for (const t of tools) {
      expect(firstSentence(t.description ?? '').length).toBeGreaterThan(5);
    }
  });
});
