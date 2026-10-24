/**
 * Exits once stdout and stderr have flushed.
 *
 * Writes to a pipe on POSIX are asynchronous: what does not fit the pipe
 * buffer (64 KiB on Linux) waits in Node's queue, and process.exit() drops
 * it, so a large --json payload read through a pipe arrives truncated.
 * Stream write callbacks run in order, so an empty write's callback means
 * everything written before it has been handed to the OS.
 */
export function exitAfterFlush(code: number): void {
  let pending = 2;
  const exit = () => {
    if (--pending === 0) process.exit(code);
  };
  process.stdout.write('', exit);
  process.stderr.write('', exit);
}
