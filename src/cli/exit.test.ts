import { execFile } from 'node:child_process';
import { expect, it } from 'vitest';

const size = 512 * 1024;

/** Runs `exit(code)` after writing `size` bytes, with stdout read through a pipe. */
function pipedOutputLength(exit: string): Promise<number> {
  const script = `import { exitAfterFlush } from './src/cli/exit.ts';
    process.stdout.write('x'.repeat(${size}));
    ${exit};`;
  return new Promise((resolve, reject) => {
    execFile(process.execPath, ['--import', 'tsx', '--input-type=module', '-e', script],
      { maxBuffer: 2 * size }, (error, stdout) => error ? reject(error) : resolve(stdout.length));
  });
}

it('keeps output larger than the pipe buffer', async () => {
  // Guard: the pipe really drops output on a bare process.exit().
  expect(await pipedOutputLength('process.exit(0)')).toBeLessThan(size);
  expect(await pipedOutputLength('exitAfterFlush(0)')).toBe(size);
});
