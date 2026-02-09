/**
 * Research script to explore Teams web app behaviour.
 * 
 * This script:
 * 1. Launches a browser with persistent context
 * 2. Navigates to Teams and handles authentication
 * 3. Monitors network requests to discover API endpoints
 * 4. Allows manual interaction to trigger searches
 * 5. Logs discovered endpoints and data structures
 */

import { chromium, type Browser, type BrowserContext, type Page, type Request, type Response } from 'playwright';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(__dirname, '../..');
const USER_DATA_DIR = path.join(PROJECT_ROOT, '.user-data');
const SESSION_STATE_PATH = path.join(PROJECT_ROOT, 'session-state.json');
const FINDINGS_PATH = path.join(PROJECT_ROOT, 'research-findings.json');

interface NetworkCapture {
  requests: CapturedRequest[];
  responses: CapturedResponse[];
}

interface CapturedRequest {
  url: string;
  method: string;
  headers: Record<string, string>;
  postData?: string;
  timestamp: string;
  resourceType: string;
}

interface CapturedResponse {
  url: string;
  status: number;
  statusText: string;
  headers: Record<string, string>;
  body?: string;
  timestamp: string;
}

// Track interesting API calls
const capturedNetworkData: NetworkCapture = {
  requests: [],
  responses: [],
};

// Patterns that indicate search-related API calls
const SEARCH_PATTERNS = [
  /search/i,
  /query/i,
  /find/i,
  /substrate/i,
  /graph\.microsoft/i,
  /teams.*api/i,
  /chatservice/i,
  /emea\.ng\.msg/i,
];

function isInterestingUrl(url: string): boolean {
  // Skip static assets, telemetry, and auth-related noise
  const skipPatterns = [
    /\.(js|css|png|jpg|jpeg|gif|svg|woff|woff2|ttf|ico)(\?|$)/i,
    /telemetry/i,
    /analytics/i,
    /logging/i,
    /beacon/i,
    /\.clarity\./i,
    /fonts\./i,
    /static/i,
  ];

  if (skipPatterns.some(pattern => pattern.test(url))) {
    return false;
  }

  // Include if it matches search patterns or is a potential API call
  return SEARCH_PATTERNS.some(pattern => pattern.test(url)) ||
    url.includes('/api/') ||
    url.includes('/v1/') ||
    url.includes('/v2/');
}

async function captureRequest(request: Request): Promise<void> {
  const url = request.url();
  
  if (!isInterestingUrl(url)) {
    return;
  }

  const captured: CapturedRequest = {
    url,
    method: request.method(),
    headers: request.headers(),
    postData: request.postData() ?? undefined,
    timestamp: new Date().toISOString(),
    resourceType: request.resourceType(),
  };

  capturedNetworkData.requests.push(captured);
  
  console.log(`\n📤 REQUEST: ${request.method()} ${url}`);
  if (captured.postData) {
    try {
      const parsed = JSON.parse(captured.postData);
      // Show full body for message sends (to see mention format)
      const isMessageSend = url.includes('/messages') && request.method() === 'POST';
      const limit = isMessageSend ? 10000 : 500;
      console.log('   Body:', JSON.stringify(parsed, null, 2).substring(0, limit));
    } catch {
      console.log('   Body:', captured.postData.substring(0, 200));
    }
  }
}

async function captureResponse(response: Response): Promise<void> {
  const url = response.url();
  
  if (!isInterestingUrl(url)) {
    return;
  }

  let body: string | undefined;
  try {
    const contentType = response.headers()['content-type'] || '';
    if (contentType.includes('application/json')) {
      body = await response.text();
    }
  } catch {
    // Response body may not be available
  }

  const captured: CapturedResponse = {
    url,
    status: response.status(),
    statusText: response.statusText(),
    headers: response.headers(),
    body,
    timestamp: new Date().toISOString(),
  };

  capturedNetworkData.responses.push(captured);

  console.log(`\n📥 RESPONSE: ${response.status()} ${url}`);
  if (body) {
    try {
      const parsed = JSON.parse(body);
      console.log('   Body preview:', JSON.stringify(parsed, null, 2).substring(0, 500));
    } catch {
      console.log('   Body preview:', body.substring(0, 200));
    }
  }
}

async function saveFindings(): Promise<void> {
  // Deduplicate and summarise findings
  const uniqueEndpoints = new Map<string, CapturedRequest>();
  
  for (const req of capturedNetworkData.requests) {
    const key = `${req.method}:${new URL(req.url).pathname}`;
    if (!uniqueEndpoints.has(key)) {
      uniqueEndpoints.set(key, req);
    }
  }

  const findings = {
    discoveredAt: new Date().toISOString(),
    uniqueEndpoints: Array.from(uniqueEndpoints.values()).map(req => ({
      method: req.method,
      url: req.url,
      hasPostData: !!req.postData,
      postDataSample: req.postData?.substring(0, 500),
    })),
    allRequests: capturedNetworkData.requests,
    allResponses: capturedNetworkData.responses,
  };

  fs.writeFileSync(FINDINGS_PATH, JSON.stringify(findings, null, 2));
  console.log(`\n💾 Findings saved to ${FINDINGS_PATH}`);
}

async function checkAuthentication(page: Page): Promise<boolean> {
  const url = page.url();
  
  // Check if we're on a login page
  if (url.includes('login.microsoftonline.com') || 
      url.includes('login.live.com') ||
      url.includes('login.microsoft.com')) {
    return false;
  }

  // Check if we're on the Teams app
  if (url.includes('teams.microsoft.com')) {
    // Look for indicators that we're logged in
    try {
      // Wait briefly for the app to load
      await page.waitForTimeout(2000);
      
      // Check for common authenticated elements
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
  console.log('\n🔐 Please log in to Microsoft Teams in the browser window...');
  console.log('   The script will continue once you are authenticated.\n');

  // Poll for authentication
  while (true) {
    if (await checkAuthentication(page)) {
      console.log('✅ Authentication detected!');
      break;
    }
    await page.waitForTimeout(2000);
  }
}

async function main(): Promise<void> {
  console.log('🔍 Teams Research Script');
  console.log('========================\n');
  console.log('This script will help discover how Teams web app works.\n');

  // Ensure user data directory exists
  if (!fs.existsSync(USER_DATA_DIR)) {
    fs.mkdirSync(USER_DATA_DIR, { recursive: true });
  }

  let browser: Browser | undefined;
  let context: BrowserContext | undefined;

  try {
    // Launch browser with persistent context
    browser = await chromium.launch({
      headless: false, // Must be visible for manual login
    });

    // Check if we have saved session state
    const hasSessionState = fs.existsSync(SESSION_STATE_PATH);

    if (hasSessionState) {
      console.log('📂 Found existing session state, attempting to restore...');
      context = await browser.newContext({
        storageState: SESSION_STATE_PATH,
        viewport: { width: 1280, height: 800 },
      });
    } else {
      console.log('📂 No session state found, starting fresh...');
      context = await browser.newContext({
        viewport: { width: 1280, height: 800 },
      });
    }

    const page = await context.newPage();

    // Set up network interception
    page.on('request', captureRequest);
    page.on('response', captureResponse);

    // Navigate to Teams
    console.log('🌐 Navigating to Teams...');
    await page.goto('https://teams.microsoft.com', { waitUntil: 'domcontentloaded' });

    // Check if authentication is needed
    const isAuthenticated = await checkAuthentication(page);
    
    if (!isAuthenticated) {
      await waitForAuthentication(page);
      
      // Save session state after successful authentication
      console.log('💾 Saving session state...');
      await context.storageState({ path: SESSION_STATE_PATH });
      console.log('✅ Session state saved!');
    } else {
      console.log('✅ Already authenticated!');
    }

    // Instructions for the user
    console.log('\n' + '='.repeat(60));
    console.log('🔬 RESEARCH MODE ACTIVE');
    console.log('='.repeat(60));
    console.log('\nThe browser is now monitoring network requests.');
    console.log('Try the following to discover API endpoints:\n');
    console.log('1. Use the search bar (Cmd/Ctrl+E or click the search icon)');
    console.log('2. Search for some text');
    console.log('3. Click on search results');
    console.log('4. Navigate between chats and channels');
    console.log('\nPress Ctrl+C in this terminal when done to save findings.');
    console.log('='.repeat(60) + '\n');

    // Keep the browser open until interrupted
    await new Promise<void>((resolve) => {
      process.on('SIGINT', () => {
        console.log('\n\n⏹️  Stopping research...');
        resolve();
      });
    });

    // Save session state before closing
    console.log('💾 Saving final session state...');
    await context.storageState({ path: SESSION_STATE_PATH });

    // Save research findings
    await saveFindings();

  } catch (error) {
    console.error('❌ Error:', error);
    throw error;
  } finally {
    if (context) {
      await context.close();
    }
    if (browser) {
      await browser.close();
    }
  }

  console.log('\n✅ Research session complete!');
  console.log(`   Review findings in: ${FINDINGS_PATH}`);
}

main().catch(console.error);
