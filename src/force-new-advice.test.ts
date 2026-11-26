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
import { recommendsForceNewWithoutApproval } from './__fixtures__/force-new-guard.js';

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

const STRING_LITERAL = /'(?:[^'\\\n]|\\.)*'|"(?:[^"\\\n]|\\.)*"|`(?:[^`\\]|\\.)*`/g;
const BARE_IDENTIFIER = /^['"`]\w+['"`]$/;

describe('forceNew advice source scan', () => {
  it('no string literal in src recommends forceNew without approval', () => {
    const offenders: string[] = [];
    for (const file of sourceFiles(SRC_DIR)) {
      const text = readFileSync(file, 'utf8');
      for (const match of text.matchAll(STRING_LITERAL)) {
        const literal = match[0];
        if (BARE_IDENTIFIER.test(literal)) continue;
        if (recommendsForceNewWithoutApproval(literal.slice(1, -1))) {
          offenders.push(`${relative(SRC_DIR, file)}: ${literal}`);
        }
      }
    }
    expect(offenders).toEqual([]);
  });
});
