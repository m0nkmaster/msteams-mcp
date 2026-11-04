/**
 * Terminal output for the `msteams` CLI.
 *
 * One generic record renderer covers every tool response: anything with an
 * array gets a numbered list, anything else gets labelled fields. Colour and
 * emoji only appear on a TTY, so piped output stays clean.
 */

const COLOUR = Boolean(process.stdout.isTTY) && !process.env.NO_COLOR;

export const bold = (text: string): string => (COLOUR ? `\x1b[1m${text}\x1b[0m` : text);
export const dim = (text: string): string => (COLOUR ? `\x1b[2m${text}\x1b[0m` : text);
export const icon = (emoji: string): string => (COLOUR ? `${emoji} ` : '');

/** Terminal width, with a sane default when output is piped. */
export const termWidth = (): number => process.stdout.columns ?? 80;

const ISO_DATE = /^\d{4}-\d{2}-\d{2}T/;

/** Fields tried, in order, as an item's headline. */
const HEADLINE_KEYS = ['content', 'displayName', 'subject', 'name', 'topic', 'title', 'text', 'preview', 'channelName', 'fileName'];

/** Response keys that are metadata rather than content. */
const META_KEYS = new Set(['success', 'pagination', 'resultCount', 'count', 'total']);

export function oneLine(text: string, max = termWidth() - 4): string {
  const flat = text.replace(/\s+/g, ' ').trim();
  return flat.length > max ? `${flat.slice(0, max - 1)}…` : flat;
}

/** Best human label for an object value (sender shapes, profiles, mentions). */
function objectLabel(value: Record<string, unknown>): string | null {
  const email = value.EmailAddress as Record<string, unknown> | undefined;
  const candidate = value.displayName ?? value.name ?? value.Name ?? email?.Name ?? email?.Address;
  return candidate === undefined || candidate === null ? null : String(candidate);
}

/** Renders any field value as a single line, or '' when there is nothing to show. */
export function summarise(value: unknown, max = termWidth() - 4, depth = 0): string {
  if (value === null || value === undefined || value === '') return '';

  if (Array.isArray(value)) {
    if (value.length === 0) return '';
    const scalars = value.filter(item => typeof item !== 'object' || item === null);
    return scalars.length === value.length
      ? oneLine(value.join(', '), max)
      : `${value.length} items`;
  }

  if (typeof value === 'object') {
    const label = objectLabel(value as Record<string, unknown>);
    if (label !== null) return label;
    if (depth > 0) return oneLine(JSON.stringify(value), max);
    // No name to show, so inline the fields: `available=true, expiresAt=...`.
    const parts = Object.entries(value as Record<string, unknown>)
      .map(([key, nested]) => [key, summarise(nested, max, depth + 1)] as const)
      .filter(([, text]) => text !== '')
      .map(([key, text]) => `${key}=${text}`);
    return parts.length ? oneLine(parts.join(', '), max) : '';
  }

  if (typeof value === 'string' && ISO_DATE.test(value)) {
    const date = new Date(value);
    if (!Number.isNaN(date.getTime())) return date.toLocaleString();
  }

  return oneLine(String(value), max);
}

/** A nested object with more fields than fit on one line, e.g. a profile. */
function isBlock(value: unknown): boolean {
  return typeof value === 'object'
    && value !== null
    && !Array.isArray(value)
    && Object.keys(value).length > 3;
}

function printFields(record: Record<string, unknown>, indent: string, skip = new Set<string>()): void {
  for (const [key, value] of Object.entries(record)) {
    if (skip.has(key) || META_KEYS.has(key)) continue;

    if (indent === '' && isBlock(value)) {
      console.log(dim(`${key}:`));
      printFields(value as Record<string, unknown>, '  ');
      continue;
    }

    const text = summarise(value, termWidth() - indent.length - key.length - 3);
    if (text) console.log(`${indent}${dim(`${key}:`)} ${text}`);
  }
}

function printItem(item: unknown, number: number): void {
  if (typeof item !== 'object' || item === null) {
    console.log(`${bold(`${number}.`)} ${summarise(item)}`);
    return;
  }

  const record = item as Record<string, unknown>;
  const headlineKey = HEADLINE_KEYS.find(key => typeof record[key] === 'string' && record[key] !== '');
  const headline = headlineKey ? oneLine(String(record[headlineKey])) : '';

  console.log(`${bold(`${number}.`)} ${headline || dim('—')}`);
  printFields(record, '   ', new Set(headlineKey ? [headlineKey] : []));
  console.log('');
}

/** Prints a successful tool response. */
export function printResponse(response: Record<string, unknown>): void {
  const listEntry = Object.entries(response)
    .find(([key, value]) => Array.isArray(value) && !META_KEYS.has(key));

  if (listEntry) {
    const [key, value] = listEntry as [string, unknown[]];
    const pagination = response.pagination as Record<string, unknown> | undefined;
    const offset = typeof pagination?.from === 'number' ? pagination.from : 0;

    const count = Number(response.resultCount ?? value.length);
    const label = count === 1 ? key.replace(/s$/, '') : key;
    console.log(`${bold(String(count))} ${label}:\n`);
    value.forEach((item, index) => printItem(item, offset + index + 1));
    printFields(response, '', new Set([key]));
  } else {
    printFields(response, '');
  }

  const pagination = response.pagination as Record<string, unknown> | undefined;
  if (pagination) {
    const total = pagination.total !== undefined ? ` of ${pagination.total}` : '';
    console.log(dim(`\nShowing ${pagination.returned}${total} (from ${pagination.from}, size ${pagination.size})`));
    if (pagination.hasMore) console.log(dim(`More available: --from ${pagination.nextFrom}`));
  }
}

/** Prints a failed tool response to stderr. */
export function printFailure(response: Record<string, unknown>): void {
  console.error(`${icon('❌')}${response.error ?? 'Unknown error'}`);
  if (response.errorCode) console.error(dim(`   code: ${response.errorCode}`));
  const suggestions = response.suggestions;
  if (Array.isArray(suggestions)) {
    for (const suggestion of suggestions) console.error(dim(`   try: ${suggestion}`));
  }
}

/** Wraps text to the terminal width with a hanging indent, clipped to maxLines. */
export function wrap(text: string, indent: number, maxLines = 2): string[] {
  const width = Math.max(40, termWidth()) - indent;
  const lines: string[] = [];
  let line = '';

  for (const word of text.replace(/\s+/g, ' ').trim().split(' ')) {
    const next = line ? `${line} ${word}` : word;
    if (next.length <= width) {
      line = next;
    } else {
      if (line) lines.push(line);
      line = word;
    }
    if (lines.length > maxLines) break;
  }
  if (line) lines.push(line);

  if (lines.length <= maxLines) return lines;
  const clipped = lines.slice(0, maxLines);
  clipped[maxLines - 1] = `${clipped[maxLines - 1].slice(0, Math.max(0, width - 2))}…`;
  return clipped;
}
