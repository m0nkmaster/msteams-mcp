/**
 * Server interface types.
 * 
 * Extracted to a shared module to avoid circular dependencies between
 * server.ts and tools/index.ts.
 */

import type { BrowserManager } from '../browser/context.js';

/**
 * Interface for the TeamsServer class.
 * 
 * This interface is used by tool handlers to interact with the server
 * without creating circular dependencies.
 */
export interface TeamsServer {
  /** Ensures a browser is running and authenticated. */
  ensureBrowser(headless?: boolean): Promise<BrowserManager>;
  
  /** Resets browser state (clears manager and initialisation flag). */
  resetBrowserState(): void;
  
  /** Gets the current browser manager, if any. */
  getBrowserManager(): BrowserManager | null;
  
  /** Sets the browser manager. */
  setBrowserManager(manager: BrowserManager): void;
  
  /** Marks the server as initialised. */
  markInitialised(): void;
  
  /** Checks if the server is initialised. */
  isInitialisedState(): boolean;
}
