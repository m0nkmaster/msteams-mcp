/**
 * Research script to understand Teams authentication redirect behaviour.
 * 
 * This script observes:
 * 1. What happens when navigating to Teams WITH a valid session
 * 2. What happens when navigating to Teams WITHOUT a session
 * 
 * Goal: Understand the fastest way to detect if we're authenticated.
 */

import { chromium, type Browser, type BrowserContext, type Page } from 'playwright';
import { hasSessionState } from '../auth/session-store.js';

const TEAMS_URL = 'https://teams.microsoft.com';

interface NavigationEvent {
  timestamp: number;
  elapsed: number;
  type: 'navigation' | 'redirect' | 'domcontentloaded' | 'load' | 'response';
  url: string;
  status?: number;
  note?: string;
}

async function observeNavigation(
  page: Page,
  scenario: string,
  startTime: number
): Promise<NavigationEvent[]> {
  const events: NavigationEvent[] = [];
  
  const log = (type: NavigationEvent['type'], url: string, extra?: Partial<NavigationEvent>) => {
    const now = Date.now();
    const event: NavigationEvent = {
      timestamp: now,
      elapsed: now - startTime,
      type,
      url,
      ...extra,
    };
    events.push(event);
    console.log(`  [${event.elapsed}ms] ${type}: ${url}${extra?.status ? ` (${extra.status})` : ''}${extra?.note ? ` - ${extra.note}` : ''}`);
  };

  // Track all frame navigations
  page.on('framenavigated', (frame) => {
    if (frame === page.mainFrame()) {
      const url = frame.url();
      const isLogin = url.includes('login.microsoftonline.com') || 
                      url.includes('login.live.com') ||
                      url.includes('login.microsoft.com');
      const isTeams = url.includes('teams.microsoft.com') ||
                      url.includes('teams.microsoft.us');
      
      log('navigation', url, { 
        note: isLogin ? 'LOGIN PAGE' : isTeams ? 'TEAMS' : undefined 
      });
    }
  });

  // Track responses (especially redirects)
  page.on('response', (response) => {
    const url = response.url();
    const status = response.status();
    
    // Only log main document responses and redirects
    if (response.request().resourceType() === 'document' || 
        (status >= 300 && status < 400)) {
      log('response', url, { 
        status,
        note: status >= 300 && status < 400 ? 'REDIRECT' : undefined
      });
    }
  });

  return events;
}

async function runScenario(
  browser: Browser,
  scenario: string,
  useSession: boolean
): Promise<{ events: NavigationEvent[]; finalUrl: string; duration: number }> {
  console.log(`\n${'='.repeat(60)}`);
  console.log(`SCENARIO: ${scenario}`);
  console.log(`Session: ${useSession ? 'WITH existing session' : 'NO session (cleared)'}`);
  console.log('='.repeat(60));

  let context: BrowserContext;
  
  if (useSession && hasSessionState()) {
    console.log('Loading existing session state...');
    // Handle encrypted session
    try {
      const { readSessionState } = await import('../auth/session-store.js');
      const state = readSessionState();
      if (state) {
        context = await browser.newContext({
          storageState: state as Parameters<typeof browser.newContext>[0] extends { storageState?: infer S } ? S : never,
          viewport: { width: 1280, height: 800 },
        });
      } else {
        throw new Error('Could not read session');
      }
    } catch {
      console.log('Could not load session, using fresh context');
      context = await browser.newContext({
        viewport: { width: 1280, height: 800 },
      });
    }
  } else {
    console.log('Using fresh context (no session)...');
    context = await browser.newContext({
      viewport: { width: 1280, height: 800 },
    });
  }

  const page = await context.newPage();
  const startTime = Date.now();
  
  console.log(`\nNavigating to ${TEAMS_URL}...`);
  console.log('Watching for navigations and redirects:\n');
  
  const events = await observeNavigation(page, scenario, startTime);
  
  try {
    // Navigate and wait for initial load
    await page.goto(TEAMS_URL, { 
      waitUntil: 'domcontentloaded',
      timeout: 30000,
    });
    
    const afterGoto = Date.now();
    console.log(`\n  [${afterGoto - startTime}ms] goto() completed (domcontentloaded)`);
    
    // Wait a bit more to see if any late redirects happen
    console.log('\n  Waiting 10 seconds to observe any delayed redirects...\n');
    
    for (let i = 1; i <= 10; i++) {
      await page.waitForTimeout(1000);
      const currentUrl = page.url();
      const isLogin = currentUrl.includes('login.microsoftonline.com') || 
                      currentUrl.includes('login.live.com') ||
                      currentUrl.includes('login.microsoft.com');
      const isTeams = currentUrl.includes('teams.microsoft.com') ||
                      currentUrl.includes('teams.microsoft.us');
      
      console.log(`  [${Date.now() - startTime}ms] Check ${i}: ${isLogin ? 'ON LOGIN' : isTeams ? 'ON TEAMS' : 'OTHER'} - ${currentUrl.substring(0, 80)}...`);
      
      // If we're on login page, we know we're not authenticated
      if (isLogin) {
        console.log('\n  ✓ Detected redirect to login page - NOT authenticated');
        break;
      }
    }
    
    const finalUrl = page.url();
    const duration = Date.now() - startTime;
    
    console.log(`\n${'─'.repeat(60)}`);
    console.log(`RESULT for "${scenario}":`);
    console.log(`  Final URL: ${finalUrl}`);
    console.log(`  Total duration: ${duration}ms`);
    console.log(`  Navigation events: ${events.length}`);
    
    const isOnLogin = finalUrl.includes('login.microsoftonline.com') || 
                      finalUrl.includes('login.live.com') ||
                      finalUrl.includes('login.microsoft.com');
    const isOnTeams = finalUrl.includes('teams.microsoft.com') ||
                      finalUrl.includes('teams.microsoft.us');
    
    console.log(`  Authenticated: ${isOnTeams && !isOnLogin ? 'YES (on Teams)' : 'NO (on login page)'}`);
    
    return { events, finalUrl, duration };
    
  } finally {
    await context.close();
  }
}

async function main(): Promise<void> {
  console.log('🔬 Teams Authentication Research');
  console.log('================================\n');
  console.log('This script observes redirect behaviour during Teams navigation.\n');

  const browser = await chromium.launch({
    headless: false, // Visible so we can see what's happening
  });

  try {
    // Scenario 1: With existing session (if available)
    if (hasSessionState()) {
      await runScenario(browser, 'WITH SESSION', true);
    } else {
      console.log('\n⚠️  No existing session found. Run `npm run cli -- login` first to create one.');
      console.log('   Skipping "WITH SESSION" scenario.\n');
    }

    // Scenario 2: Without session (cleared context)
    await runScenario(browser, 'WITHOUT SESSION', false);

    console.log('\n\n' + '='.repeat(60));
    console.log('SUMMARY');
    console.log('='.repeat(60));
    console.log(`
Key observations to look for:
1. How quickly does the redirect to login.microsoftonline.com happen?
2. Does the redirect happen BEFORE or AFTER domcontentloaded?
3. Are there any intermediate URLs?
4. What HTTP status codes are used for redirects?

Based on this, we can tune LOGIN_REDIRECT_TIMEOUT_MS in auth.ts
`);

  } finally {
    await browser.close();
  }
}

main().catch(console.error);
