/**
 * Command and argument resolution for the `msteams` CLI.
 *
 * Everything here is derived from the MCP tool definitions - their names and
 * JSON input schemas - so a new tool is callable, documented and type-coerced
 * without touching the CLI.
 */

import { parseArgs } from 'node:util';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';

/** Thrown for anything the user can fix by retyping the command (exit code 2). */
export class UsageError extends Error {}

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
  return `Flags for '${label}': ${flags}\nRun 'msteams help ${label}' for details.`;
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
export function parseToolArgs(tool: Tool, argv: string[], label = tool.name): ParsedToolArgs {
  const props = schemaProps(tool);

  const options: Record<string, { type: 'boolean' | 'string'; short?: string }> = { ...GLOBAL_FLAGS };
  const optionType = (prop?: SchemaProp): 'boolean' | 'string' =>
    prop?.type === 'boolean' ? 'boolean' : 'string';

  for (const [name, prop] of Object.entries(props)) {
    options[name] = { type: optionType(prop) };
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
    if (flag === 'json' || flag === 'help') continue;
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
