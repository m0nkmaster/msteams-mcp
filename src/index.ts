#!/usr/bin/env node
/**
 * Teams MCP Server entry point.
 * 
 * This MCP server enables AI assistants to interact with Microsoft Teams
 * via direct API calls — searching messages, sending replies, managing
 * favourites, calendar, files, and more. The browser is only used for
 * initial authentication.
 */

import { runServer } from './server.js';

runServer().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});
