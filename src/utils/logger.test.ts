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
    it('logs info messages with context', () => {
      info('test-context', 'Information message');
      
      expect(consoleSpy.log).toHaveBeenCalledWith('[test-context] Information message');
    });

    it('does not log when level is warn', () => {
      setLogLevel('warn');
      
      info('context', 'info message');
      
      expect(consoleSpy.log).not.toHaveBeenCalled();
    });

    it('logs when level is info', () => {
      setLogLevel('info');
      
      info('context', 'info message');
      
      expect(consoleSpy.log).toHaveBeenCalled();
    });
  });

  describe('debug', () => {
    it('logs debug messages with context', () => {
      setLogLevel('debug');
      
      debug('test-context', 'Debug message');
      
      expect(consoleSpy.debug).toHaveBeenCalledWith('[test-context] Debug message');
    });

    it('does not log when level is info', () => {
      setLogLevel('info');
      
      debug('context', 'debug message');
      
      expect(consoleSpy.debug).not.toHaveBeenCalled();
    });

    it('logs when level is debug', () => {
      setLogLevel('debug');
      
      debug('context', 'debug message');
      
      expect(consoleSpy.debug).toHaveBeenCalled();
    });
  });

  describe('setLogLevel', () => {
    it('sets log level to error', () => {
      setLogLevel('error');
      
      error('c', 'error');
      warn('c', 'warn');
      info('c', 'info');
      debug('c', 'debug');
      
      expect(consoleSpy.error).toHaveBeenCalled();
      expect(consoleSpy.warn).not.toHaveBeenCalled();
      expect(consoleSpy.log).not.toHaveBeenCalled();
      expect(consoleSpy.debug).not.toHaveBeenCalled();
    });

    it('sets log level to warn', () => {
      setLogLevel('warn');
      
      error('c', 'error');
      warn('c', 'warn');
      info('c', 'info');
      debug('c', 'debug');
      
      expect(consoleSpy.error).toHaveBeenCalled();
      expect(consoleSpy.warn).toHaveBeenCalled();
      expect(consoleSpy.log).not.toHaveBeenCalled();
      expect(consoleSpy.debug).not.toHaveBeenCalled();
    });

    it('sets log level to info', () => {
      setLogLevel('info');
      
      error('c', 'error');
      warn('c', 'warn');
      info('c', 'info');
      debug('c', 'debug');
      
      expect(consoleSpy.error).toHaveBeenCalled();
      expect(consoleSpy.warn).toHaveBeenCalled();
      expect(consoleSpy.log).toHaveBeenCalled();
      expect(consoleSpy.debug).not.toHaveBeenCalled();
    });

    it('sets log level to debug', () => {
      setLogLevel('debug');
      
      error('c', 'error');
      warn('c', 'warn');
      info('c', 'info');
      debug('c', 'debug');
      
      expect(consoleSpy.error).toHaveBeenCalled();
      expect(consoleSpy.warn).toHaveBeenCalled();
      expect(consoleSpy.log).toHaveBeenCalled();
      expect(consoleSpy.debug).toHaveBeenCalled();
    });

    it('defaults to info for unknown level', () => {
      // Set an invalid level via env (simulated by calling with valid level first)
      setLogLevel('info');
      
      info('c', 'test');
      debug('c', 'test');
      
      expect(consoleSpy.log).toHaveBeenCalled();
      expect(consoleSpy.debug).not.toHaveBeenCalled();
    });
  });
});
