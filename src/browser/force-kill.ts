/**
 * Force-kills browser processes still holding the MCP browser profile.
 *
 * Used when a persistent-context browser cannot be closed cleanly (wedged
 * renderer/GPU process). Only processes whose command line references our
 * profile directory are touched — the user's own browser uses a different
 * profile and is never matched.
 */

import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import * as log from '../utils/logger.js';

const execFileAsync = promisify(execFile);

/** Per-process listing/killing budget (ms). */
const KILL_TIMEOUT_MS = 15000;

/**
 * Kills every Edge/Chrome process launched with `profileDir`.
 * Returns the number of processes killed. Never throws.
 */
export async function killBrowserProfileProcesses(profileDir: string): Promise<number> {
  try {
    const killed = process.platform === 'win32'
      ? await killWindows(profileDir)
      : await killPosix(profileDir);
    if (killed > 0) {
      log.warn('browser', `Force-killed ${killed} wedged browser process(es) holding the profile`);
    }
    return killed;
  } catch (error) {
    log.warn('browser', `Could not clean up wedged browser processes: ${error instanceof Error ? error.message : String(error)}`);
    return 0;
  }
}

/** Windows: match on the process command line via CIM, then taskkill /T. */
async function killWindows(profileDir: string): Promise<number> {
  const script =
    `Get-CimInstance Win32_Process -Filter "Name='msedge.exe' OR Name='chrome.exe'"` +
    ` | Select-Object ProcessId,CommandLine | ConvertTo-Json -Compress`;
  const { stdout } = await execFileAsync(
    'powershell', ['-NoProfile', '-NonInteractive', '-Command', script],
    { timeout: KILL_TIMEOUT_MS },
  );

  const trimmed = stdout.trim();
  if (!trimmed) return 0;
  const parsed: unknown = JSON.parse(trimmed);
  const processes = (Array.isArray(parsed) ? parsed : [parsed]) as Array<{
    ProcessId?: unknown;
    CommandLine?: unknown;
  }>;

  const pids = processes
    .filter(p => typeof p?.ProcessId === 'number'
      && typeof p.CommandLine === 'string'
      && p.CommandLine.includes(profileDir))
    .map(p => p.ProcessId as number);

  for (const pid of pids) {
    // /T takes the whole process tree: renderers may lack the profile arg.
    await execFileAsync('taskkill', ['/PID', String(pid), '/F', '/T'], { timeout: KILL_TIMEOUT_MS })
      .catch(() => {});
  }
  return pids.length;
}

/** POSIX: pgrep -f matches the full command line, then SIGKILL each pid. */
async function killPosix(profileDir: string): Promise<number> {
  const { stdout } = await execFileAsync('pgrep', ['-f', profileDir], { timeout: KILL_TIMEOUT_MS })
    .catch(() => ({ stdout: '' }));
  const pids = stdout
    .split('\n')
    .map(line => parseInt(line.trim(), 10))
    .filter(pid => Number.isFinite(pid) && pid > 0 && pid !== process.pid);

  for (const pid of pids) {
    try {
      process.kill(pid, 'SIGKILL');
    } catch {
      // Process already gone.
    }
  }
  return pids.length;
}
