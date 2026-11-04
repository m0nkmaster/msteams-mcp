/**
 * Unit tests for logger utilities.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { error, warn, info, debug, setLogLevel } from './logger.js';

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
});
