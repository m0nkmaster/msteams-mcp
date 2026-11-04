/**
 * Help text for the `msteams` CLI, generated from the tool definitions.
 *
 * The index stays to one screen: one line per command, first sentence only.
 * Full parameter detail lives in `msteams help <command>`, and the complete
 * schemas in `msteams help --json`.
 */

import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import { FLAG_ALIASES, positionalProp, requiredProps, schemaProps, shortName } from './resolve.js';
import { bold, dim, termWidth, wrap } from './render.js';

/**
 * First sentence of a description. The tool descriptions are written for an
 * LLM and run long, so only the opening sentence is useful in a command list.
 * A period followed by a non-capital (as in "e.g.") does not end a sentence.
 */
export function firstSentence(text: string): string {
  const match = /^(.*?[.!?])(?=\s+[A-Z]|$)/s.exec(text.trim());
  const sentence = (match?.[1] ?? text).replace(/\s+/g, ' ').trim();
  return sentence.replace(/\.$/, '');
}

function flagLabel(prop: { type?: string; enum?: string[] }): string {
  if (prop.enum) return `<${prop.enum.join('|')}>`;
  if (prop.type === 'boolean') return '';
  if (prop.type === 'array') return '<a,b>';
  return `<${prop.type ?? 'string'}>`;
}

/** Aliases pointing at a parameter, e.g. conversationId -> ['--to']. */
function aliasesFor(prop: string): string[] {
  return Object.entries(FLAG_ALIASES)
    .filter(([, target]) => target === prop)
    .map(([alias]) => `--${alias}`);
}

/** The one-screen command index. */
export function renderIndex(tools: Tool[], groups: Array<{ name: string; tools: string[] }>, version: string): string {
  const names = tools.map(tool => tool.name);
  const byName = new Map(tools.map(tool => [tool.name, tool]));
  const labels = new Map(names.map(name => [name, shortName(names, name)]));
  const column = Math.min(22, Math.max(...[...labels.values()].map(label => label.length)) + 2);

  const lines: string[] = [
    `${bold('msteams')} ${dim(version)} - Microsoft Teams from your terminal`,
    '',
    bold('Usage'),
    '  msteams <command> [value] [--flag value]',
    `  msteams help <command>    ${dim('flags and detail for one command')}`,
    `  msteams <command> --json  ${dim('machine-readable output')}`,
  ];

  for (const group of groups) {
    const members = group.tools.filter(name => byName.has(name));
    if (members.length === 0) continue;
    lines.push('', bold(group.name));
    for (const name of members) {
      const label = labels.get(name) ?? name;
      const summary = firstSentence(byName.get(name)?.description ?? '');
      const room = termWidth() - column - 4;
      lines.push(`  ${label.padEnd(column)}${dim(summary.length > room ? `${summary.slice(0, room - 1)}…` : summary)}`);
    }
  }

  lines.push(
    '',
    dim('Commands are tool names without the teams_/get_ prefix; any unambiguous'),
    dim("abbreviation works (send -> send_message). 'msteams login' first."),
  );
  return lines.join('\n');
}

/** Detail for one command: what it does, how to call it, every flag. */
export function renderToolHelp(tools: Tool[], tool: Tool): string {
  const names = tools.map(t => t.name);
  const label = shortName(names, tool.name);
  const props = schemaProps(tool);
  const required = new Set(requiredProps(tool));
  const positional = positionalProp(tool);

  const usage = [`  msteams ${label}`];
  if (positional) usage[0] += ` <${positional}>`;
  if (Object.keys(props).length > (positional ? 1 : 0)) usage[0] += ' [--flag value]';

  const lines: string[] = [
    `${bold(`msteams ${label}`)} ${dim(`(${tool.name})`)}`,
    '',
    ...wrap(firstSentence(tool.description ?? ''), 0, 3).map(line => `  ${line}`),
    '',
    bold('Usage'),
    ...usage,
  ];

  const entries = Object.entries(props);
  if (entries.length > 0) {
    lines.push('', bold('Flags'));
    const column = Math.min(28, Math.max(...entries.map(([name, prop]) => `--${name} ${flagLabel(prop)}`.trim().length)) + 2);

    for (const [name, prop] of entries) {
      const flag = `--${name} ${flagLabel(prop)}`.trim();
      const alias = aliasesFor(name);
      const detail = [
        required.has(name) ? 'Required.' : '',
        alias.length ? `(or ${alias.join(', ')})` : '',
        prop.description ?? '',
      ].filter(Boolean).join(' ');
      const wrapped = wrap(detail, column + 2, 2);

      // A flag too long for the column takes a line of its own.
      if (flag.length >= column) {
        lines.push(`  ${bold(flag)}`);
        for (const line of wrapped) lines.push(`  ${' '.repeat(column)}${line}`);
        continue;
      }
      lines.push(`  ${bold(flag.padEnd(column))}${wrapped[0] ?? ''}`);
      for (const line of wrapped.slice(1)) lines.push(`  ${' '.repeat(column)}${line}`);
    }
  }

  lines.push('', dim('Add --json for the raw payload.'));
  return lines.join('\n');
}
