#!/usr/bin/env npx tsx
/**
 * MCP Protocol Test Harness
 * 
 * Tests the MCP server by connecting a client through the actual MCP protocol,
 * rather than calling underlying functions directly. This ensures the full
 * protocol layer works correctly.
 * 
 * Usage:
 *   npm run test:mcp                              # List tools and check status
 *   npm run test:mcp -- search "query"            # Search for messages (shortcut)
 *   npm run test:mcp -- teams_search --query "q"  # Generic tool call
 *   npm run test:mcp -- --json                    # Output as JSON
 * 
 * Any unrecognised command is treated as a tool name. Use --key value for parameters.
 */

import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { createServer } from '../server.js';

// Shortcuts map command names to tool names and parameter mappings
const SHORTCUTS: Record<string, { tool: string; primaryArg?: string }> = {
  search: { tool: 'teams_search', primaryArg: 'query' },
  status: { tool: 'teams_status' },
  send: { tool: 'teams_send_message', primaryArg: 'content' },
  me: { tool: 'teams_get_me' },
  people: { tool: 'teams_search_people', primaryArg: 'query' },
  favorites: { tool: 'teams_get_favorites' },
  save: { tool: 'teams_save_message' },
  unsave: { tool: 'teams_unsave_message' },
  thread: { tool: 'teams_get_thread' },
  login: { tool: 'teams_login' },
  contacts: { tool: 'teams_get_frequent_contacts' },
  channel: { tool: 'teams_find_channel', primaryArg: 'query' },
  chat: { tool: 'teams_get_chat', primaryArg: 'userId' },
  groupchat: { tool: 'teams_create_group_chat' },
  unread: { tool: 'teams_get_unread' },
  markread: { tool: 'teams_mark_read' },
  activity: { tool: 'teams_get_activity' },
};

// Map CLI flags to tool parameter names
const FLAG_MAPPINGS: Record<string, string> = {
  '--to': 'conversationId',
  '--message': 'messageId',
  '--reply': 'replyToMessageId',
  '--replyTo': 'replyToMessageId',
  '--from': 'from',
  '--size': 'size',
  '--limit': 'limit',
  '--query': 'query',
  '--content': 'content',
  '--force': 'forceNew',
  '--user': 'userId',
  '--userId': 'userId',
  '--userIds': 'userIds',
  '--topic': 'topic',
  '--markRead': 'markRead',
};

interface ParsedArgs {
  command: string;  // 'list' or tool name
  toolName: string | null;
  primaryArg: string | null;
  args: Record<string, unknown>;
  json: boolean;
}

function parseArgs(): ParsedArgs {
  const args = process.argv.slice(2);
  const result: ParsedArgs = {
    command: 'list',
    toolName: null,
    primaryArg: null,
    args: {},
    json: false,
  };

  let i = 0;
  
  // First pass: find the command (first non-flag argument)
  while (i < args.length) {
    const arg = args[i];
    
    if (arg === '--json') {
      result.json = true;
      i++;
      continue;
    }
    
    if (arg.startsWith('--')) {
      // Skip flag and its value
      i += 2;
      continue;
    }
    
    // Found the command
    result.command = arg;
    i++;
    break;
  }

  // Check if it's a shortcut
  const shortcut = SHORTCUTS[result.command];
  if (shortcut) {
    result.toolName = shortcut.tool;
    
    // Look for primary argument (next non-flag arg)
    while (i < args.length) {
      const arg = args[i];
      if (!arg.startsWith('--')) {
        if (shortcut.primaryArg) {
          result.args[shortcut.primaryArg] = arg;
        }
        result.primaryArg = arg;
        i++;
        break;
      }
      i++;
    }
  } else if (result.command !== 'list') {
    // Treat as a tool name (add teams_ prefix if not present)
    result.toolName = result.command.startsWith('teams_') 
      ? result.command 
      : `teams_${result.command}`;
  }

  // Second pass: collect all --key value pairs
  for (let j = 0; j < args.length; j++) {
    const arg = args[j];
    
    if (arg === '--json') {
      result.json = true;
      continue;
    }
    
    if (arg.startsWith('--') && args[j + 1] !== undefined) {
      const key = FLAG_MAPPINGS[arg] || arg.slice(2);  // Use mapping or strip --
      let value: unknown = args[j + 1];
      
      // Only parse booleans and specific numeric fields
      // Don't coerce messageId, conversationId etc. - they're strings
      const numericFields = new Set(['from', 'size', 'limit']);
      if (value === 'true') value = true;
      else if (value === 'false') value = false;
      else if (numericFields.has(key) && /^\d+$/.test(value as string)) {
        value = parseInt(value as string, 10);
      } else if (typeof value === 'string' && (value.startsWith('[') || value.startsWith('{'))) {
        // Try to parse JSON for array or object values
        try {
          value = JSON.parse(value);
        } catch {
          // Keep as string if not valid JSON
        }
      }
      
      result.args[key] = value;
      j++;  // Skip the value
    }
  }

  return result;
}

function log(message: string): void {
  console.log(message);
}

function logSection(title: string): void {
  console.log('\n' + '─'.repeat(50));
  console.log(`  ${title}`);
  console.log('─'.repeat(50));
}

async function createTestClient(): Promise<{ client: Client; cleanup: () => Promise<void> }> {
  // Create the MCP server
  const server = await createServer();
  
  // Create linked in-memory transports
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  
  // Connect the server to its transport
  await server.connect(serverTransport);
  
  // Create and connect the client
  const client = new Client(
    { name: 'mcp-test-harness', version: '1.0.0' },
    { capabilities: {} }
  );
  
  await client.connect(clientTransport);
  
  const cleanup = async () => {
    await client.close();
    await server.close();
  };
  
  return { client, cleanup };
}

async function listTools(client: Client, json: boolean): Promise<void> {
  if (!json) {
    logSection('Available Tools');
  }
  
  const result = await client.listTools();
  
  if (json) {
    console.log(JSON.stringify(result, null, 2));
    return;
  }
  
  log(`Found ${result.tools.length} tools:\n`);
  
  for (const tool of result.tools) {
    log(`📦 ${tool.name}`);
    log(`   ${tool.description}`);
    
    const schema = tool.inputSchema;
    if (schema.properties) {
      const props = Object.entries(schema.properties);
      const required = new Set(schema.required ?? []);
      
      for (const [name, prop] of props) {
        const propObj = prop as { type?: string; description?: string };
        const reqMark = required.has(name) ? ' (required)' : '';
        log(`   - ${name}: ${propObj.type ?? 'any'}${reqMark}`);
        if (propObj.description) {
          log(`     ${propObj.description}`);
        }
      }
    }
    log('');
  }
  
  // Show available shortcuts
  log('Shortcuts:');
  for (const [shortcut, config] of Object.entries(SHORTCUTS)) {
    const primaryNote = config.primaryArg ? ` <${config.primaryArg}>` : '';
    log(`  ${shortcut}${primaryNote} → ${config.tool}`);
  }
  log('');
  log('Any unrecognised command is treated as a tool name.');
  log('Use --key value for parameters, e.g.: teams_find_channel --query "name"');
}

async function callTool(
  client: Client,
  toolName: string,
  args: Record<string, unknown>,
  json: boolean
): Promise<void> {
  // Verify the tool exists
  const tools = await client.listTools();
  const tool = tools.tools.find(t => t.name === toolName);
  
  if (!tool) {
    console.error(`❌ Unknown tool: ${toolName}`);
    console.error('');
    console.error('Available tools:');
    for (const t of tools.tools) {
      console.error(`  - ${t.name}`);
    }
    process.exit(1);
  }

  if (!json) {
    logSection(`Calling: ${toolName}`);
    if (Object.keys(args).length > 0) {
      log(`Arguments: ${JSON.stringify(args)}\n`);
    }
  }

  // Login tool needs a longer timeout for MFA flows
  const timeout = toolName === 'teams_login' ? 5 * 60 * 1000 : undefined; // 5 minutes for login
  const result = await client.callTool({ name: toolName, arguments: args }, undefined, { timeout });

  if (json) {
    console.log(JSON.stringify(result, null, 2));
    return;
  }

  // Pretty-print the result
  const content = result.content as Array<{ type: string; text?: string }>;
  const textContent = content.find(c => c.type === 'text');

  if (textContent?.text) {
    try {
      const response = JSON.parse(textContent.text);
      prettyPrintResponse(response, toolName);
    } catch {
      // Not JSON, just print as-is
      log(textContent.text);
    }
  } else {
    log('(No text content in response)');
  }
}

/**
 * Pretty-prints a tool response based on common patterns
 */
function prettyPrintResponse(response: Record<string, unknown>, _toolName: string): void {
  // Check for error
  if (response.success === false) {
    log(`❌ Failed: ${response.error || 'Unknown error'}`);
    if (response.code) log(`   Code: ${response.code}`);
    if (response.suggestion) log(`   Suggestion: ${response.suggestion}`);
    return;
  }

  log('✅ Success\n');

  // Handle common response shapes
  if (response.results && Array.isArray(response.results)) {
    printResultsList(response.results, response);
  } else if (response.favorites && Array.isArray(response.favorites)) {
    printFavoritesList(response.favorites);
  } else if (response.messages && Array.isArray(response.messages)) {
    printMessagesList(response.messages);
  } else if (response.activities && Array.isArray(response.activities)) {
    printActivityList(response.activities);
  } else if (response.contacts && Array.isArray(response.contacts)) {
    printContactsList(response.contacts);
  } else if (response.profile) {
    printProfile(response.profile as Record<string, unknown>);
  } else {
    // Generic output for other responses
    printGenericResponse(response);
  }

  // Print pagination if present
  if (response.pagination) {
    const p = response.pagination as Record<string, unknown>;
    log(`\nPagination: from=${p.from}, size=${p.size}, returned=${p.returned}`);
    if (p.total !== undefined) log(`Total available: ${p.total}`);
    if (p.hasMore) log(`More results available (use --from ${p.nextFrom})`);
  }
}

function printResultsList(results: unknown[], response: Record<string, unknown>): void {
  log(`Found ${response.resultCount ?? results.length} results:\n`);
  
  for (let i = 0; i < results.length; i++) {
    const r = results[i] as Record<string, unknown>;
    const num = ((response.pagination as Record<string, unknown>)?.from as number ?? 0) + i + 1;
    
    // Try to extract meaningful content
    const content = String(r.content ?? r.displayName ?? r.name ?? '').substring(0, 100).replace(/\n/g, ' ');
    log(`${num}. ${content}${content.length >= 100 ? '...' : ''}`);
    
    // Print common fields
    const sender = extractSenderName(r.sender);
    if (sender) log(`   From: ${sender}`);
    if (r.email) log(`   Email: ${r.email}`);
    if (r.teamName) log(`   Team: ${r.teamName}`);
    if (r.channelName && r.channelName !== r.teamName) log(`   Channel: ${r.channelName}`);
    if (r.timestamp) log(`   Time: ${r.timestamp}`);
    if (r.conversationId) log(`   ConversationId: ${r.conversationId}`);
    if (r.jobTitle) log(`   Title: ${r.jobTitle}`);
    if (r.department) log(`   Dept: ${r.department}`);
    if (r.mri) log(`   MRI: ${r.mri}`);
    log('');
  }
}

function printFavoritesList(favorites: unknown[]): void {
  log(`Found ${favorites.length} favourites:\n`);
  
  for (const f of favorites) {
    const fav = f as Record<string, unknown>;
    const typeLabel = fav.conversationType ? ` [${fav.conversationType}]` : '';
    const nameLabel = fav.displayName || '(unnamed)';
    log(`⭐ ${nameLabel}${typeLabel}`);
    log(`   ID: ${fav.conversationId}`);
  }
}

function printMessagesList(messages: unknown[]): void {
  log(`Got ${messages.length} messages:\n`);
  
  for (const msg of messages) {
    const m = msg as Record<string, unknown>;
    const preview = String(m.content ?? '').substring(0, 100).replace(/\n/g, ' ');
    const sender = (m.sender as Record<string, unknown>)?.displayName || 
                   (m.sender as Record<string, unknown>)?.mri || 'Unknown';
    const time = m.timestamp ? new Date(m.timestamp as string).toLocaleString() : '';
    const fromMe = m.isFromMe ? ' (you)' : '';
    
    log(`📝 ${sender}${fromMe} - ${time}`);
    log(`   ${preview}${String(m.content ?? '').length > 100 ? '...' : ''}`);
    log('');
  }
}

function printContactsList(contacts: unknown[]): void {
  log(`Found ${contacts.length} contacts:\n`);
  
  for (const c of contacts) {
    const contact = c as Record<string, unknown>;
    log(`👤 ${contact.displayName}`);
    if (contact.email) log(`   Email: ${contact.email}`);
    if (contact.jobTitle) log(`   Title: ${contact.jobTitle}`);
    if (contact.department) log(`   Dept: ${contact.department}`);
    log('');
  }
}

function printActivityList(activities: unknown[]): void {
  log(`Found ${activities.length} activity items:\n`);
  
  const typeIcons: Record<string, string> = {
    mention: '📣',
    reaction: '👍',
    reply: '💬',
    message: '📝',
    unknown: '❓',
  };
  
  for (const a of activities) {
    const activity = a as Record<string, unknown>;
    const type = activity.type as string || 'unknown';
    const icon = typeIcons[type] || '❓';
    const sender = (activity.sender as Record<string, unknown>)?.displayName || 'Unknown';
    const time = activity.timestamp ? new Date(activity.timestamp as string).toLocaleString() : '';
    const topic = activity.topic ? ` in "${activity.topic}"` : '';
    const preview = String(activity.content ?? '').substring(0, 80).replace(/\n/g, ' ');
    
    log(`${icon} [${type}] ${sender}${topic} - ${time}`);
    log(`   ${preview}${String(activity.content ?? '').length > 80 ? '...' : ''}`);
    if (activity.conversationId) log(`   ConversationId: ${activity.conversationId}`);
    log('');
  }
}

function printProfile(profile: Record<string, unknown>): void {
  log('👤 Profile:\n');
  for (const [key, value] of Object.entries(profile)) {
    if (value !== null && value !== undefined) {
      log(`   ${key}: ${value}`);
    }
  }
}

function printGenericResponse(response: Record<string, unknown>): void {
  // Filter out success flag and print remaining fields
  for (const [key, value] of Object.entries(response)) {
    if (key === 'success') continue;
    
    if (typeof value === 'object' && value !== null) {
      log(`${key}: ${JSON.stringify(value, null, 2)}`);
    } else {
      log(`${key}: ${value}`);
    }
  }
}

function extractSenderName(sender: unknown): string | null {
  if (!sender) return null;
  if (typeof sender === 'string') return sender;
  if (typeof sender === 'object') {
    const s = sender as Record<string, unknown>;
    // Handle { EmailAddress: { Name: string } } structure
    if (s.EmailAddress && typeof s.EmailAddress === 'object') {
      const email = s.EmailAddress as Record<string, unknown>;
      if (email.Name) return String(email.Name);
      if (email.Address) return String(email.Address);
    }
    // Handle { name: string } or { displayName: string } structure
    if (s.displayName) return String(s.displayName);
    if (s.name) return String(s.name);
    if (s.Name) return String(s.Name);
  }
  return null;
}

async function main(): Promise<void> {
  const parsed = parseArgs();
  
  if (!parsed.json) {
    console.log('\n🧪 MCP Protocol Test Harness');
    console.log('============================');
  }
  
  let cleanup: (() => Promise<void>) | null = null;
  
  try {
    const { client, cleanup: cleanupFn } = await createTestClient();
    cleanup = cleanupFn;
    
    if (!parsed.json) {
      log('\n✅ Connected to MCP server via in-memory transport');
    }
    
    if (parsed.command === 'list' || !parsed.toolName) {
      await listTools(client, parsed.json);
    } else {
      await callTool(client, parsed.toolName, parsed.args, parsed.json);
    }
    
    if (!parsed.json) {
      logSection('Complete');
      log('MCP protocol test finished successfully.');
    }
    
  } catch (error) {
    console.error('\n❌ Error:', error instanceof Error ? error.message : String(error));
    if (error instanceof Error && error.stack) {
      console.error('\nStack trace:');
      console.error(error.stack);
    }
    process.exit(1);
  } finally {
    if (cleanup) {
      await cleanup();
    }
  }
}

main();
