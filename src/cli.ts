#!/usr/bin/env node
/**
 * Teams CLI
 *
 * A standalone command-line interface with full parity to the MCP server. It
 * runs the same in-process server over an in-memory transport and speaks to it
 * through the real MCP protocol, so every registered tool is callable.
 *
 *   msteams                          # command list
 *   msteams help search              # detail for one command
 *   msteams search "query"           # call a tool
 *   msteams get_thread --to "<id>"   # flags come from the tool's schema
 *   msteams search "query" --json    # machine-readable output
 *
 * Commands, flags, type coercion and help are all derived from the tool
 * definitions, so a new tool needs no changes here.
 *
 * Exit codes: 0 success, 1 the tool failed, 2 the command was malformed.
 */

import { createRequire } from 'module';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import { createServer } from './server.js';
import { toolCategories } from './tools/registry.js';
import { UsageError, parseToolArgs, resolveToolName, shortName } from './cli/resolve.js';
import { printFailure, printResponse } from './cli/render.js';
import { renderIndex, renderToolHelp } from './cli/help.js';

const require = createRequire(import.meta.url);
const pkg = require('../package.json') as { version: string };

/** Login can wait on SSO/MFA in a browser; everything else uses the default. */
const LOGIN_TIMEOUT_MS = 5 * 60 * 1000;

async function connect(): Promise<{ client: Client; cleanup: () => Promise<void> }> {
  const server = await createServer();
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await server.connect(serverTransport);

  const client = new Client({ name: 'msteams', version: pkg.version }, { capabilities: {} });
  await client.connect(clientTransport);

  return {
    client,
    cleanup: async () => {
      await client.close();
      await server.close();
    },
  };
}

/** Extracts the tool payload from an MCP response, or null if it is not JSON. */
function toolPayload(result: unknown): Record<string, unknown> | null {
  const content = (result as { content?: Array<{ type: string; text?: string }> }).content ?? [];
  const text = content.find(part => part.type === 'text')?.text;
  if (!text) return null;
  try {
    return JSON.parse(text) as Record<string, unknown>;
  } catch {
    console.log(text);
    return null;
  }
}

async function run(argv: string[]): Promise<number> {
  if (argv.includes('--version') || argv[0] === '-v') {
    console.log(pkg.version);
    return 0;
  }

  const commandIndex = argv.findIndex(arg => !arg.startsWith('-'));
  const command = commandIndex === -1 ? null : argv[commandIndex];
  const rest = argv.filter((_, index) => index !== commandIndex);

  const { client, cleanup } = await connect();
  try {
    const { tools } = await client.listTools();
    const names = tools.map(tool => tool.name);

    // Help: bare invocation, `help`, `help <command>` or `<command> --help`.
    const wantsHelp = command === null || command === 'help';
    if (wantsHelp) {
      const target = rest.find(arg => !arg.startsWith('-'));
      const tool = target ? tools.find(t => t.name === resolveToolName(names, target)) : undefined;
      if (rest.includes('--json')) {
        console.log(JSON.stringify(tool ?? { tools }, null, 2));
      } else {
        console.log(tool ? renderToolHelp(tools, tool) : renderIndex(tools, toolCategories, pkg.version));
      }
      return 0;
    }

    const name = resolveToolName(names, command);
    const tool = tools.find(t => t.name === name) as Tool;
    const label = shortName(names, name);
    const { args, json, help } = parseToolArgs(tool, rest, label);

    if (help) {
      console.log(renderToolHelp(tools, tool));
      return 0;
    }

    const timeout = name === 'teams_login' ? LOGIN_TIMEOUT_MS : undefined;
    const result = await client.callTool({ name, arguments: args }, undefined, { timeout });
    const payload = toolPayload(result);
    if (!payload) return result.isError ? 1 : 0;

    if (json) {
      console.log(JSON.stringify(payload, null, 2));
    } else if (payload.success === false) {
      printFailure(payload);
    } else {
      printResponse(payload);
    }

    return payload.success === false ? 1 : 0;
  } finally {
    await cleanup();
  }
}

run(process.argv.slice(2))
  .then(code => process.exit(code))
  .catch((error: unknown) => {
    const message = error instanceof Error ? error.message : String(error);
    console.error(message);
    if (!(error instanceof UsageError) && error instanceof Error && process.env.LOG_LEVEL === 'debug') {
      console.error(error.stack);
    }
    process.exit(error instanceof UsageError ? 2 : 1);
  });
