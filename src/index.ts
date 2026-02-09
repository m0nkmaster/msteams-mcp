#!/usr/bin/env node
/**
 * Teams MCP Server entry point.
 * 
 * This MCP server enables AI assistants to search Microsoft Teams
 * messages using browser automation.
 */

import { runServer } from './server.js';

runServer().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});
