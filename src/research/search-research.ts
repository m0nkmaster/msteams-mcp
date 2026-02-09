/**
 * Focused research script for Teams search API.
 * 
 * Captures full request/response bodies for Substrate search endpoints
 * to understand pagination, query format, and response structure.
 * 
 * Usage:
 *   npm run research:search
 * 
 * Instructions:
 *   1. Run this script
 *   2. Perform a search in Teams
 *   3. Click "Messages" tab or "See more results"
 *   4. Scroll through results to trigger pagination
 *   5. Press Ctrl+C to stop and save findings
 */

import { chromium, type Browser, type BrowserContext, type Page, type Request, type Response } from 'playwright';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(__dirname, '../..');
const SESSION_STATE_PATH = path.join(PROJECT_ROOT, 'session-state.json');
const SEARCH_FINDINGS_PATH = path.join(PROJECT_ROOT, 'search-api-findings.json');

interface SearchCapture {
  timestamp: string;
  type: 'request' | 'response';
  url: string;
  method?: string;
  status?: number;
  body?: unknown;
  headers?: Record<string, string>;
}

const captures: SearchCapture[] = [];

// Focus specifically on search-related endpoints
function isSearchEndpoint(url: string): boolean {
  const searchPatterns = [
    /substrate\.office\.com\/search/i,
    /\/api\/v\d+\/query/i,
    /\/api\/v\d+\/suggestions/i,
    /\/search\//i,
  ];
  return searchPatterns.some(p => p.test(url));
}

async function captureSearchRequest(request: Request): Promise<void> {
  const url = request.url();
  if (!isSearchEndpoint(url)) return;

  const postData = request.postData();
  let body: unknown;
  
  if (postData) {
    try {
      body = JSON.parse(postData);
    } catch {
      body = postData;
    }
  }

  const capture: SearchCapture = {
    timestamp: new Date().toISOString(),
    type: 'request',
    url,
    method: request.method(),
    body,
    headers: request.headers(),
  };
  
  captures.push(capture);

  console.log('\n' + '━'.repeat(80));
  console.log(`📤 SEARCH REQUEST: ${request.method()} ${url}`);
  console.log('━'.repeat(80));
  
  if (body) {
    console.log('\n📋 Request Body:');
    console.log(JSON.stringify(body, null, 2));
  }
  
  // Highlight pagination-related fields
  if (body && typeof body === 'object') {
    const bodyObj = body as Record<string, unknown>;
    const paginationFields = ['Size', 'From', 'Skip', 'Top', 'Offset', 'PageSize', 'Page', 'MaxResults'];
    const foundPagination: Record<string, unknown> = {};
    
    function findPaginationFields(obj: unknown, prefix = ''): void {
      if (!obj || typeof obj !== 'object') return;
      
      for (const [key, value] of Object.entries(obj as Record<string, unknown>)) {
        const fullKey = prefix ? `${prefix}.${key}` : key;
        if (paginationFields.some(p => key.toLowerCase().includes(p.toLowerCase()))) {
          foundPagination[fullKey] = value;
        }
        if (typeof value === 'object') {
          findPaginationFields(value, fullKey);
        }
      }
    }
    
    findPaginationFields(bodyObj);
    
    if (Object.keys(foundPagination).length > 0) {
      console.log('\n📊 Pagination Fields Found:');
      console.log(JSON.stringify(foundPagination, null, 2));
    }
  }
}

async function captureSearchResponse(response: Response): Promise<void> {
  const url = response.url();
  if (!isSearchEndpoint(url)) return;

  let body: unknown;
  try {
    const contentType = response.headers()['content-type'] || '';
    if (contentType.includes('application/json')) {
      const text = await response.text();
      body = JSON.parse(text);
    }
  } catch {
    // Response body may not be available
  }

  const capture: SearchCapture = {
    timestamp: new Date().toISOString(),
    type: 'response',
    url,
    status: response.status(),
    body,
  };
  
  captures.push(capture);

  console.log('\n' + '━'.repeat(80));
  console.log(`📥 SEARCH RESPONSE: ${response.status()} ${url}`);
  console.log('━'.repeat(80));
  
  if (body && typeof body === 'object') {
    const bodyObj = body as Record<string, unknown>;
    
    // Show structure summary
    console.log('\n📋 Response Structure:');
    const summarise = (obj: unknown, depth = 0): string => {
      if (depth > 2) return '...';
      if (Array.isArray(obj)) {
        return `Array[${obj.length}]${obj.length > 0 ? ` of ${summarise(obj[0], depth + 1)}` : ''}`;
      }
      if (obj && typeof obj === 'object') {
        const keys = Object.keys(obj as Record<string, unknown>);
        if (keys.length > 8) {
          return `{${keys.slice(0, 5).join(', ')}, ... +${keys.length - 5} more}`;
        }
        return `{${keys.join(', ')}}`;
      }
      return typeof obj;
    };
    console.log(summarise(bodyObj));
    
    // Show pagination-related response fields
    const paginationResponseFields = ['Total', 'TotalCount', 'Count', 'HasMore', 'HasNext', 'NextLink', 'SkipToken', 'ContinuationToken'];
    const foundPagination: Record<string, unknown> = {};
    
    function findPaginationFields(obj: unknown, prefix = ''): void {
      if (!obj || typeof obj !== 'object') return;
      
      for (const [key, value] of Object.entries(obj as Record<string, unknown>)) {
        const fullKey = prefix ? `${prefix}.${key}` : key;
        if (paginationResponseFields.some(p => key.toLowerCase().includes(p.toLowerCase()))) {
          foundPagination[fullKey] = value;
        }
        if (typeof value === 'object' && !Array.isArray(value)) {
          findPaginationFields(value, fullKey);
        }
      }
    }
    
    findPaginationFields(bodyObj);
    
    if (Object.keys(foundPagination).length > 0) {
      console.log('\n📊 Pagination Response Fields:');
      console.log(JSON.stringify(foundPagination, null, 2));
    }
    
    // Show results count
    const countArrays = (obj: unknown): { path: string; count: number }[] => {
      const results: { path: string; count: number }[] = [];
      
      function traverse(o: unknown, path: string): void {
        if (Array.isArray(o) && o.length > 0) {
          results.push({ path: path || 'root', count: o.length });
        }
        if (o && typeof o === 'object' && !Array.isArray(o)) {
          for (const [key, value] of Object.entries(o as Record<string, unknown>)) {
            traverse(value, path ? `${path}.${key}` : key);
          }
        }
      }
      
      traverse(obj, '');
      return results;
    };
    
    const arrays = countArrays(bodyObj);
    if (arrays.length > 0) {
      console.log('\n📈 Result Counts:');
      arrays.forEach(a => console.log(`   ${a.path}: ${a.count} items`));
    }
    
    // Show first result sample if available
    const findFirstResult = (obj: unknown): unknown => {
      if (Array.isArray(obj) && obj.length > 0) {
        const first = obj[0];
        if (first && typeof first === 'object' && Object.keys(first as Record<string, unknown>).length > 2) {
          return first;
        }
      }
      if (obj && typeof obj === 'object') {
        for (const value of Object.values(obj as Record<string, unknown>)) {
          const found = findFirstResult(value);
          if (found) return found;
        }
      }
      return null;
    };
    
    const firstResult = findFirstResult(bodyObj);
    if (firstResult) {
      console.log('\n📄 First Result Sample:');
      console.log(JSON.stringify(firstResult, null, 2).substring(0, 1000));
    }
  }
}

async function checkAuthentication(page: Page): Promise<boolean> {
  const url = page.url();
  
  if (url.includes('login.microsoftonline.com') || 
      url.includes('login.live.com') ||
      url.includes('login.microsoft.com')) {
    return false;
  }

  if (url.includes('teams.microsoft.com')) {
    try {
      await page.waitForTimeout(2000);
      const hasAppBar = await page.locator('[data-tid="app-bar"]').count() > 0;
      const hasSearchBox = await page.locator('[data-tid="search-box"]').count() > 0 ||
                          await page.locator('input[placeholder*="Search"]').count() > 0;
      return hasAppBar || hasSearchBox;
    } catch {
      return false;
    }
  }

  return false;
}

async function waitForAuthentication(page: Page): Promise<void> {
  console.log('\n🔐 Please log in to Microsoft Teams...');

  while (true) {
    if (await checkAuthentication(page)) {
      console.log('✅ Authentication detected!');
      break;
    }
    await page.waitForTimeout(2000);
  }
}

async function saveFindings(): Promise<void> {
  const findings = {
    capturedAt: new Date().toISOString(),
    totalCaptures: captures.length,
    requests: captures.filter(c => c.type === 'request'),
    responses: captures.filter(c => c.type === 'response'),
  };

  fs.writeFileSync(SEARCH_FINDINGS_PATH, JSON.stringify(findings, null, 2));
  console.log(`\n💾 Search findings saved to ${SEARCH_FINDINGS_PATH}`);
}

async function main(): Promise<void> {
  console.log('🔍 Teams Search API Research');
  console.log('============================\n');

  let browser: Browser | undefined;
  let context: BrowserContext | undefined;

  try {
    browser = await chromium.launch({ headless: false });

    const hasSessionState = fs.existsSync(SESSION_STATE_PATH);

    if (hasSessionState) {
      console.log('📂 Restoring session state...');
      context = await browser.newContext({
        storageState: SESSION_STATE_PATH,
        viewport: { width: 1400, height: 900 },
      });
    } else {
      console.log('📂 No session state found, starting fresh...');
      context = await browser.newContext({
        viewport: { width: 1400, height: 900 },
      });
    }

    const page = await context.newPage();

    // Set up network interception for search endpoints only
    page.on('request', captureSearchRequest);
    page.on('response', captureSearchResponse);

    console.log('🌐 Navigating to Teams...');
    await page.goto('https://teams.microsoft.com', { waitUntil: 'domcontentloaded' });

    const isAuthenticated = await checkAuthentication(page);
    
    if (!isAuthenticated) {
      await waitForAuthentication(page);
      await context.storageState({ path: SESSION_STATE_PATH });
      console.log('✅ Session state saved!');
    } else {
      console.log('✅ Already authenticated!');
    }

    console.log('\n' + '═'.repeat(60));
    console.log('🔬 SEARCH RESEARCH MODE');
    console.log('═'.repeat(60));
    console.log('\nTo capture search API data:');
    console.log('');
    console.log('  1. Press Cmd/Ctrl+E to open search');
    console.log('  2. Type a search query and press Enter');
    console.log('  3. Click "Messages" tab to see message results');
    console.log('  4. Click "See all results" if available');
    console.log('  5. Scroll through results to trigger pagination');
    console.log('');
    console.log('Watch for 📊 Pagination Fields in the output!');
    console.log('');
    console.log('Press Ctrl+C when done to save findings.');
    console.log('═'.repeat(60) + '\n');

    await new Promise<void>((resolve) => {
      process.on('SIGINT', () => {
        console.log('\n\n⏹️  Stopping research...');
        resolve();
      });
    });

    await context.storageState({ path: SESSION_STATE_PATH });
    await saveFindings();

  } catch (error) {
    console.error('❌ Error:', error);
    throw error;
  } finally {
    if (context) await context.close();
    if (browser) await browser.close();
  }

  console.log('\n✅ Search research complete!');
  console.log(`   Review findings in: ${SEARCH_FINDINGS_PATH}`);
}

main().catch(console.error);
