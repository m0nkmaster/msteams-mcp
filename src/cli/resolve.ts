/**
 * Command and argument resolution for the `msteams` CLI.
 *
 * Everything here is derived from the MCP tool definitions - their names and
 * JSON input schemas - so a new tool is callable, documented and type-coerced
 * without touching the CLI.
 */

import { readFileSync, readSync, realpathSync, statSync } from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { isatty } from 'node:tty';
import { parseArgs } from 'node:util';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import { CONFIG_DIR } from '../auth/session-store.js';
import {
  CREDENTIAL_FILE_BASENAMES,
  forbiddenRoots,
  relativeInside,
  windowsLexicalRefusal,
} from '../api/chat-attachments-api.js';
import { MAX_SIMPLE_UPLOAD_BYTES } from '../constants.js';

/** Thrown for anything the user can fix by retyping the command (exit code 2). */
export class UsageError extends Error {}

/** Schema property that takes the message text, and the flag that reads it from a file. */
export const CONTENT_PROP = 'content';
export const CONTENT_FILE_FLAG = 'content-file';

/** The subset of JSON Schema the tool definitions use. */
export interface SchemaProp {
  type?: string;
  description?: string;
  enum?: string[];
  items?: { type?: string };
}

/** Flag aliases that read better than the raw parameter name. */
export const FLAG_ALIASES: Record<string, string> = {
  to: 'conversationId',
  msg: 'messageId',
  message: 'messageId',
  reply: 'replyToMessageId',
  replyTo: 'replyToMessageId',
  user: 'userId',
  force: 'forceNew',
};

/** Flags accepted for every command. */
const GLOBAL_FLAGS: Record<string, { type: 'boolean' | 'string'; short?: string }> = {
  json: { type: 'boolean' },
  help: { type: 'boolean', short: 'h' },
};

export function schemaProps(tool: Tool): Record<string, SchemaProp> {
  return (tool.inputSchema.properties ?? {}) as Record<string, SchemaProp>;
}

export function requiredProps(tool: Tool): string[] {
  return (tool.inputSchema.required ?? []) as string[];
}

/** The parameter a bare value after the command maps to: the first required string. */
export function positionalProp(tool: Tool): string | null {
  const props = schemaProps(tool);
  return requiredProps(tool).find(name => (props[name]?.type ?? 'string') === 'string') ?? null;
}

const normalise = (name: string): string => name.replace(/_/g, '').toLowerCase();

/**
 * Resolves a typed command to a tool name.
 *
 * Tries, in order: the exact name, `teams_<cmd>`, `teams_get_<cmd>`,
 * `teams_search_<cmd>`, a unique `teams_<cmd>*` prefix, then a unique
 * substring match ignoring underscores. So `search`, `get_thread`, `thread`,
 * `send` and `markread` all land on the right tool.
 */
export function resolveToolName(names: string[], command: string): string {
  const cmd = command.toLowerCase();

  const exact = [cmd, `teams_${cmd}`, `teams_get_${cmd}`, `teams_search_${cmd}`]
    .find(name => names.includes(name));
  if (exact) return exact;

  const prefixed = names.filter(name => name.startsWith(`teams_${cmd}`));
  if (prefixed.length === 1) return prefixed[0];

  const fuzzy = names.filter(name => normalise(name).includes(normalise(cmd)));
  if (fuzzy.length === 1) return fuzzy[0];

  const candidates = prefixed.length ? prefixed : fuzzy;
  throw new UsageError(
    candidates.length
      ? `Ambiguous command '${command}' - matches ${candidates.join(', ')}.`
      : `Unknown command '${command}'. Run 'msteams help' for the command list.`
  );
}

/** The shortest form that still resolves back to this tool - used in help output. */
export function shortName(names: string[], toolName: string): string {
  const candidates = [
    toolName.replace(/^teams_(get_|search_|list_|find_)?/, ''),
    toolName.replace(/^teams_/, ''),
  ];
  for (const candidate of candidates) {
    if (!candidate) continue;
    try {
      if (resolveToolName(names, candidate) === toolName) return candidate;
    } catch {
      // Ambiguous - fall through to the next candidate.
    }
  }
  return toolName;
}

/** Coerces one raw flag value using the schema's declared type. */
function coerce(prop: SchemaProp | undefined, raw: string, flag: string): unknown {
  const type = prop?.type ?? 'string';

  if (type === 'number' || type === 'integer') {
    const value = Number(raw);
    if (!Number.isFinite(value)) throw new UsageError(`--${flag} expects a number, got '${raw}'`);
    return value;
  }

  if (type === 'array') {
    const items = raw.trim().startsWith('[')
      ? parseJson(raw, flag)
      : raw.split(',').map(part => part.trim()).filter(Boolean);
    return prop?.items?.type === 'number' ? (items as unknown[]).map(Number) : items;
  }

  if (type === 'object') return parseJson(raw, flag);

  return raw;
}

function parseJson(raw: string, flag: string): unknown {
  try {
    return JSON.parse(raw);
  } catch {
    throw new UsageError(`--${flag} expects JSON, got '${raw}'`);
  }
}

/** One-line reminder of a command's flags, used in every usage error. */
function flagHint(props: Record<string, SchemaProp>, label: string): string {
  const flags = Object.keys(props).map(name => `--${name}`).join(' ') || '(none)';
  const contentFile = props[CONTENT_PROP]?.type === 'string' ? ` --${CONTENT_FILE_FLAG} <path|->` : '';
  return `Flags for '${label}': ${flags}${contentFile}\nRun 'msteams help ${label}' for details.`;
}

/**
 * Decodes --content-file bytes to text. UTF-16 files (what Windows
 * PowerShell 5.1 writes by default) are decoded by their BOM; anything else
 * must be valid UTF-8 - mojibake sent to Teams is worse than an error.
 */
function decodeContentFile(buf: Buffer): string {
  let encoding = 'utf-8';
  let bytes = buf;
  if (buf.length >= 2 && buf[0] === 0xFF && buf[1] === 0xFE) {
    encoding = 'utf-16le';
    bytes = buf.subarray(2);
  } else if (buf.length >= 2 && buf[0] === 0xFE && buf[1] === 0xFF) {
    encoding = 'utf-16be';
    bytes = buf.subarray(2);
  }
  try {
    return new TextDecoder(encoding, { fatal: true }).decode(bytes);
  } catch {
    throw new Error(`file is not valid ${encoding === 'utf-8' ? 'UTF-8' : 'UTF-16'} - save it as UTF-8`);
  }
}

/** Realpath when possible, plain resolve otherwise (protected dirs may not exist). */
function resolveRealLocal(p: string): string {
  try {
    // .native resolves Windows 8.3 short names that plain realpathSync keeps.
    return realpathSync.native(p);
  } catch {
    return path.resolve(p);
  }
}

/**
 * Path guard for --content-file, mirroring the --files upload guard minus the
 * upload-root containment (message text may live anywhere the user points).
 * Returns the refusal reason, or null when the path may be read. Checks run
 * on the realpath so junctions and 8.3 aliases cannot slip past them.
 */
export function contentFilePathRefusal(
  candidate: string,
  forbiddenDirs: string[] = forbiddenRoots(process.platform, os.homedir(), process.env, CONFIG_DIR),
  maxBytes: number = MAX_SIMPLE_UPLOAD_BYTES,
  homeDir: string = os.homedir(),
): string | null {
  // Windows lexical traps before any fs call: UNC/device paths would open
  // an SMB session on realpath/stat and leak NTLM.
  if (process.platform === 'win32') {
    const refusal = windowsLexicalRefusal(candidate, path.parse(candidate).root);
    if (refusal) return `--content-file path ${refusal}: ${candidate}`;
  }

  let realPath: string;
  try {
    // .native resolves Windows 8.3 short names that plain realpathSync keeps.
    realPath = realpathSync.native(candidate);
  } catch {
    return null; // unreadable - the reader reports "Cannot read ..."
  }

  // Dot-paths are where credentials live. Inside the home directory only
  // direct children are refused (~/.ssh, ~/.aws, ~/.config, ~/.env) so
  // project dot-dirs like .tmp or .worktrees stay usable; outside home
  // every dot segment is refused. A dot-file basename (.env, .npmrc) is
  // refused anywhere.
  const relHome = relativeInside(realPath, resolveRealLocal(homeDir));
  const dotSegments = relHome === null
    ? realPath.split(/[\\/]/)
    : relHome.split(/[\\/]/).slice(0, 1);
  const basename = path.basename(realPath);
  if (basename.startsWith('.') ||
      dotSegments.some(s => s.startsWith('.') && s !== '.' && s !== '..')) {
    return `--content-file path must not contain dotfiles or dot-directories: ${candidate}`;
  }
  if (CREDENTIAL_FILE_BASENAMES.test(path.basename(realPath))) {
    return `--content-file is a credential file and cannot be sent as a message: ${candidate}`;
  }
  for (const forbidden of forbiddenDirs) {
    if (relativeInside(realPath, resolveRealLocal(forbidden)) !== null) {
      return `--content-file path is inside a protected directory: ${candidate}`;
    }
  }
  try {
    const stat = statSync(realPath);
    if (!stat.isFile()) return `--content-file is not a regular file: ${candidate}`;
    if (stat.size > maxBytes) {
      return `--content-file exceeds the ${maxBytes / (1024 * 1024)} MB limit: ${candidate}`;
    }
  } catch {
    return `--content-file could not be read: ${candidate}`;
  }
  return null;
}

/** Reads stdin up to maxBytes; throws UsageError past the cap. */
function readStdinCapped(maxBytes: number): Buffer {
  const chunks: Buffer[] = [];
  const chunk = Buffer.allocUnsafe(64 * 1024);
  let total = 0;
  for (;;) {
    const n = readSync(0, chunk, 0, chunk.length, null);
    if (n === 0) break;
    total += n;
    if (total > maxBytes) {
      throw new UsageError(`--content-file - exceeds the ${maxBytes / (1024 * 1024)} MB limit.`);
    }
    chunks.push(Buffer.from(chunk.subarray(0, n)));
  }
  return Buffer.concat(chunks);
}

/** Reads the --content-file value: a path, or '-' for stdin (fd 0). */
const defaultReadContentFile = (pathOrDash: string): string => {
  if (pathOrDash === '-') {
    if (isatty(0)) {
      throw new UsageError('--content-file - reads stdin; pipe the message in or pass a file path');
    }
    return decodeContentFile(readStdinCapped(MAX_SIMPLE_UPLOAD_BYTES));
  }
  // `npm run cli` runs the script with cwd = the package root while the
  // caller's directory survives in INIT_CWD - but only trust INIT_CWD inside
  // an npm run; a directly-invoked CLI (the documented bypass) uses cwd.
  const base = process.env.npm_lifecycle_event
    ? (process.env.INIT_CWD ?? process.cwd())
    : process.cwd();
  const target = path.resolve(base, pathOrDash);
  const refusal = contentFilePathRefusal(target);
  if (refusal) throw new UsageError(refusal);
  return decodeContentFile(readFileSync(target));
};

/**
 * Normalises file text for a message: strip a leading UTF-8 BOM, convert
 * CRLF to LF, and drop trailing newlines only - other whitespace is content.
 */
function normalizeContentFileText(raw: string): string {
  return raw
    .replace(/^\uFEFF/, '')
    .replace(/\r\n/g, '\n')
    .replace(/\n+$/, '');
}


export interface ParsedToolArgs {
  args: Record<string, unknown>;
  json: boolean;
  help: boolean;
}

/**
 * Parses a command's arguments against its input schema.
 *
 * Flags are whatever the schema declares (plus the aliases above), values are
 * coerced to the schema's types, and an unknown flag is a usage error rather
 * than something the server has to reject.
 */
export function parseToolArgs(
  tool: Tool,
  argv: string[],
  label = tool.name,
  readContent: (pathOrDash: string) => string = defaultReadContentFile
): ParsedToolArgs {
  const props = schemaProps(tool);

  const options: Record<string, { type: 'boolean' | 'string'; short?: string }> = { ...GLOBAL_FLAGS };
  const optionType = (prop?: SchemaProp): 'boolean' | 'string' =>
    prop?.type === 'boolean' ? 'boolean' : 'string';

  for (const [name, prop] of Object.entries(props)) {
    options[name] = { type: optionType(prop) };
  }
  // --content-file exists only where the schema has a string `content`
  // property; elsewhere it stays an unknown-flag usage error.
  if (props[CONTENT_PROP]?.type === 'string') {
    options[CONTENT_FILE_FLAG] = { type: 'string' };
  }
  for (const [alias, target] of Object.entries(FLAG_ALIASES)) {
    if (props[target] && !options[alias]) options[alias] = { type: optionType(props[target]) };
  }

  // `--flag true|false` stays supported for boolean flags (it is how the docs
  // and the MCP skill spell them), so take the value here rather than letting
  // parseArgs treat it as a positional.
  const explicitBooleans: Record<string, boolean> = {};
  const tokens: string[] = [];
  for (let i = 0; i < argv.length; i++) {
    const token = argv[i];
    const name = token.startsWith('--') ? token.slice(2) : '';
    const next = argv[i + 1];
    if (options[name]?.type === 'boolean' && (next === 'true' || next === 'false')) {
      explicitBooleans[FLAG_ALIASES[name] ?? name] = next === 'true';
      i++;
      continue;
    }
    tokens.push(token);
  }

  let parsed: ReturnType<typeof parseArgs>;
  try {
    parsed = parseArgs({ args: tokens, options, allowPositionals: true, strict: true });
  } catch (error) {
    // parseArgs messages carry a long tail about '--'; the first sentence is the useful part.
    const reason = (error instanceof Error ? error.message : String(error)).split('. ')[0];
    throw new UsageError(`${reason}.\n${flagHint(props, label)}`);
  }

  // `--help` only wants to read about the command, so skip validation.
  if (parsed.values.help) {
    return { args: {}, json: Boolean(parsed.values.json), help: true };
  }

  const args: Record<string, unknown> = {};
  for (const [flag, raw] of Object.entries(parsed.values)) {
    if (flag === 'json' || flag === 'help' || flag === CONTENT_FILE_FLAG) continue;
    const key = FLAG_ALIASES[flag] ?? flag;
    args[key] = typeof raw === 'boolean' ? raw : coerce(props[key], String(raw), flag);
  }
  Object.assign(args, explicitBooleans);

  const missing = requiredProps(tool).filter(name => args[name] === undefined);

  const positional = positionalProp(tool);
  const [value, ...extra] = parsed.positionals;
  if (extra.length) {
    throw new UsageError(`Unexpected argument '${extra[0]}'. Run 'msteams help ${label}'.`);
  }
  if (value !== undefined) {
    if (!positional) {
      throw new UsageError(`'${label}' takes no value. Run 'msteams help ${label}' for its flags.`);
    }
    if (args[positional] === undefined) {
      args[positional] = coerce(props[positional], value, positional);
      missing.splice(missing.indexOf(positional), 1);
    }
  }

  const contentFileArg = parsed.values[CONTENT_FILE_FLAG];
  if (contentFileArg !== undefined) {
    if (args[CONTENT_PROP] !== undefined) {
      throw new UsageError('give the message as text or --content-file, not both');
    }
    let raw: string;
    try {
      raw = readContent(String(contentFileArg));
    } catch (error) {
      if (error instanceof UsageError) throw error;
      const reason = error instanceof Error ? error.message : String(error);
      throw new UsageError(`Cannot read --content-file '${contentFileArg}': ${reason}`);
    }
    const text = normalizeContentFileText(raw);
    if (text.includes('\u0000')) {
      throw new UsageError(`--content-file '${contentFileArg}' contains NUL bytes - it is probably UTF-16 without a BOM; save it as UTF-8.`);
    }
    if (!text.trim()) {
      throw new UsageError(`--content-file '${contentFileArg}' is empty or only whitespace.`);
    }
    args[CONTENT_PROP] = text;
    const contentIndex = missing.indexOf(CONTENT_PROP);
    if (contentIndex !== -1) missing.splice(contentIndex, 1);
  }

  if (missing.length > 0) {
    const positionalNote = positional && missing.includes(positional)
      ? ` (${positional} can be given as the first value: msteams ${label} "...")`
      : '';
    throw new UsageError(
      `Missing required ${missing.map(name => `--${name}`).join(', ')}${positionalNote}.\n` +
      flagHint(props, label)
    );
  }

  return {
    args,
    json: Boolean(parsed.values.json),
    help: Boolean(parsed.values.help),
  };
}
