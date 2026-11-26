/**
 * Source scan: no string literal in src/ may recommend forceNew without
 * user approval. Complements the per-error-code sweep in errors.test.ts —
 * this catches advice embedded anywhere else (tool descriptions, API error
 * messages, CLI text).
 */

import { describe, expect, it } from 'vitest';
import { readdirSync, readFileSync } from 'node:fs';
import { join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  recommendsForceNewWithoutApproval,
  stringLiteralValues,
} from './__fixtures__/force-new-guard.js';

const SRC_DIR = fileURLToPath(new URL('.', import.meta.url));

function* sourceFiles(dir: string): Generator<string> {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) {
      if (entry.name !== '__fixtures__') yield* sourceFiles(full);
    } else if (entry.name.endsWith('.ts') && !entry.name.endsWith('.test.ts')) {
      yield full;
    }
  }
}

describe('stringLiteralValues', () => {
  it('ignores regex literals and extracts real string values', () => {
    const text = [
      'const re = /`[^`]+`/;',
      "const s = 'call teams_login with forceNew: true';",
      'const t = "double quoted";',
      "const f = 'forceNew';",
    ].join('\n');
    expect(stringLiteralValues('x.ts', text)).toEqual([
      'call teams_login with forceNew: true',
      'double quoted',
    ]);
  });

  it('returns unescaped values', () => {
    const text = "const s = 'Only with the user\\'s approval, call teams_login';";
    expect(stringLiteralValues('x.ts', text)).toEqual([
      "Only with the user's approval, call teams_login",
    ]);
  });

  it('joins a template expression into one value', () => {
    const text = "const s = `Call teams_login with ${'forceNew'}: true`;";
    expect(stringLiteralValues('x.ts', text)).toEqual([
      'Call teams_login with forceNew: true',
    ]);
  });

  it('joins a +-concatenation containing a string literal', () => {
    const text = "const s = 'call ' + name + ' teams_login with forceNew: true';";
    expect(stringLiteralValues('x.ts', text)).toEqual([
      'call … teams_login with forceNew: true',
    ]);
  });

  it('inlines \\w+ literals inside concatenations but skips standalone ones', () => {
    const text = "const a = 'forceNew'; const b = 'go ' + 'forceNew';";
    expect(stringLiteralValues('x.ts', text)).toEqual(['go forceNew']);
  });
});

describe('recommendsForceNewWithoutApproval', () => {
  it('flags an approval phrase negated in the same clause', () => {
    expect(recommendsForceNewWithoutApproval(
      'Do not ask the user to approve; call teams_login with forceNew: true',
    )).toBe(true);
  });

  it('does not count approval on a different clause, line, or after a colon', () => {
    expect(recommendsForceNewWithoutApproval(
      'Ask the user to approve\ncall teams_login with forceNew: true',
    )).toBe(true);
    expect(recommendsForceNewWithoutApproval(
      'Ask the user to approve: call teams_login with forceNew: true',
    )).toBe(true);
  });

  it('flags force = true and force: true mentions', () => {
    expect(recommendsForceNewWithoutApproval('Call teams_login with force = true')).toBe(true);
    expect(recommendsForceNewWithoutApproval('Call teams_login with force: true')).toBe(true);
  });

  it('does not flag MUST together with without forceNew', () => {
    expect(recommendsForceNewWithoutApproval('You MUST call teams_login without forceNew')).toBe(false);
  });
});

describe('forceNew advice source scan', () => {
  it('no string literal in src recommends forceNew without approval', () => {
    const offenders: string[] = [];
    for (const file of sourceFiles(SRC_DIR)) {
      for (const value of stringLiteralValues(file, readFileSync(file, 'utf8'))) {
        if (recommendsForceNewWithoutApproval(value)) {
          offenders.push(`${relative(SRC_DIR, file)}: '${value}'`);
        }
      }
    }
    expect(offenders).toEqual([]);
  });
});
