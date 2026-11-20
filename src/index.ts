#!/usr/bin/env node
/**
 * Teams MCP Server entry point.
 * 
 * This MCP server enables AI assistants to interact with Microsoft Teams
 * via direct API calls — searching messages, sending replies, managing
 * favourites, calendar, files, and more. The browser is only used for
 * initial authentication.
 */

import * as path from 'path';
import { runServer } from './server.js';
import { CONFIG_DIR } from './auth/session-store.js';
import { setLogFilePath } from './utils/logger.js';

// MCP clients do not always retain the server's stderr; keep a durable copy.
setLogFilePath(path.join(CONFIG_DIR, 'logs', 'server.log'));

runServer().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});
