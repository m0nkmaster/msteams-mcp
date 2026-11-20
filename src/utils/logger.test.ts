/**
 * Unit tests for logger utilities.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { error, warn, info, debug, setLogLevel, setLogFilePath } from './logger.js';

describe('logger', () => {
  let consoleSpy: {
    error: ReturnType<typeof vi.spyOn>;
    warn: ReturnType<typeof vi.spyOn>;
    log: ReturnType<typeof vi.spyOn>;
    debug: ReturnType<typeof vi.spyOn>;
  };

  beforeEach(() => {
    consoleSpy = {
      error: vi.spyOn(console, 'error').mockImplementation(() => {}),
      warn: vi.spyOn(console, 'warn').mockImplementation(() => {}),
      log: vi.spyOn(console, 'log').mockImplementation(() => {}),
      debug: vi.spyOn(console, 'debug').mockImplementation(() => {}),
    };
    // Reset to default log level before each test
    setLogLevel('info');
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  describe('error', () => {
    it('logs error messages with context', () => {
      error('test-context', 'Something went wrong');
      
      expect(consoleSpy.error).toHaveBeenCalledWith('[test-context] Something went wrong');
    });

    it('logs even when level is above error', () => {
      setLogLevel('error');
      
      error('context', 'error message');
      
      expect(consoleSpy.error).toHaveBeenCalled();
    });
  });

  describe('warn', () => {
    it('logs warning messages with context', () => {
      warn('test-context', 'This is a warning');
      
      expect(consoleSpy.warn).toHaveBeenCalledWith('[test-context] This is a warning');
    });

    it('does not log when level is error', () => {
      setLogLevel('error');
      
      warn('context', 'warning');
      
      expect(consoleSpy.warn).not.toHaveBeenCalled();
    });

    it('logs when level is warn', () => {
      setLogLevel('warn');
      
      warn('context', 'warning');
      
      expect(consoleSpy.warn).toHaveBeenCalled();
    });
  });

  describe('info', () => {
    it('logs info messages with context, on stderr', () => {
      info('test-context', 'Information message');
      
      expect(consoleSpy.error).toHaveBeenCalledWith('[test-context] Information message');
    });

    it('does not log when level is warn', () => {
      setLogLevel('warn');
      
      info('context', 'info message');
      
      expect(consoleSpy.error).not.toHaveBeenCalled();
    });

    it('logs when level is info', () => {
      setLogLevel('info');
      
      info('context', 'info message');
      
      expect(consoleSpy.error).toHaveBeenCalled();
    });
  });

  describe('debug', () => {
    it('logs debug messages with context, on stderr', () => {
      setLogLevel('debug');
      
      debug('test-context', 'Debug message');
      
      expect(consoleSpy.error).toHaveBeenCalledWith('[test-context] Debug message');
    });

    it('does not log when level is info', () => {
      setLogLevel('info');
      
      debug('context', 'debug message');
      
      expect(consoleSpy.error).not.toHaveBeenCalled();
    });

    it('logs when level is debug', () => {
      setLogLevel('debug');
      
      debug('context', 'debug message');
      
      expect(consoleSpy.error).toHaveBeenCalled();
    });
  });

  describe('setLogLevel', () => {
    const emitAllLevels = () => {
      error('c', 'error');
      warn('c', 'warn');
      info('c', 'info');
      debug('c', 'debug');
    };

    // error, info and debug all write to stderr, so count the messages that
    // got through rather than which console method carried them.
    it.each([
      ['error', 1, 0],
      ['warn', 1, 1],
      ['info', 2, 1],
      ['debug', 3, 1],
    ] as const)('level %s lets %i through to stderr', (level, stderrCalls, warnCalls) => {
      setLogLevel(level);

      emitAllLevels();

      expect(consoleSpy.error).toHaveBeenCalledTimes(stderrCalls);
      expect(consoleSpy.warn).toHaveBeenCalledTimes(warnCalls);
    });

    // stdout is the MCP stdio protocol channel and the CLI's data output, so
    // nothing may log there.
    it('never writes to stdout', () => {
      setLogLevel('debug');

      emitAllLevels();

      expect(consoleSpy.log).not.toHaveBeenCalled();
      expect(consoleSpy.debug).not.toHaveBeenCalled();
    });
  });

  describe('file logging', () => {
    let dir: string;
    let logFile: string;

    beforeEach(() => {
      dir = fs.mkdtempSync(path.join(os.tmpdir(), 'mcp-logger-'));
      logFile = path.join(dir, 'logs', 'server.log');
      setLogFilePath(logFile);
    });

    afterEach(() => {
      setLogFilePath(null);
      fs.rmSync(dir, { recursive: true, force: true });
    });

    it('appends ISO-timestamped lines to the log file', () => {
      info('ctx', 'hello file');

      const content = fs.readFileSync(logFile, 'utf8');
      expect(content).toMatch(/^\d{4}-\d{2}-\d{2}T[\d:.]+Z \[ctx\] hello file\n$/m);
    });

    it('respects the configured level for file output', () => {
      setLogLevel('error');

      info('ctx', 'filtered out');
      error('ctx', 'kept');

      const content = fs.readFileSync(logFile, 'utf8');
      expect(content).not.toContain('filtered out');
      expect(content).toContain('[ctx] kept');
    });

    it('rotates the file past 1 MB, keeping one backup', () => {
      fs.mkdirSync(path.dirname(logFile), { recursive: true });
      fs.writeFileSync(logFile, 'x'.repeat(1024 * 1024 + 10));

      info('ctx', 'after-rotate');

      expect(fs.statSync(logFile).size).toBeLessThan(1024);
      expect(fs.statSync(`${logFile}.1`).size).toBeGreaterThan(1024 * 1024);
      expect(fs.readFileSync(logFile, 'utf8')).toContain('after-rotate');
    });

    it('never throws when the file cannot be written', () => {
      const blocker = path.join(dir, 'blocker');
      fs.writeFileSync(blocker, 'x');
      setLogFilePath(path.join(blocker, 'logs', 'server.log'));

      expect(() => info('ctx', 'still works')).not.toThrow();
      expect(consoleSpy.error).toHaveBeenCalledWith('[ctx] still works');
    });
  });
});
