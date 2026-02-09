#!/usr/bin/env npx tsx
/**
 * Debug script to inspect the Teams search page.
 * Takes screenshots and dumps page structure for debugging selectors.
 * 
 * Usage:
 *   npm run debug:search
 *   npm run debug:search -- "search query"
 */

import { createBrowserContext, closeBrowser } from '../browser/context.js';
import { ensureAuthenticated } from '../browser/auth.js';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(__dirname, '../..');
const DEBUG_DIR = path.join(PROJECT_ROOT, 'debug-output');

async function main(): Promise<void> {
  const query = process.argv[2] ?? 'test';
  
  console.log('🔍 Debug Search Script\n');
  console.log(`Query: "${query}"`);
  
  // Ensure debug output directory exists
  if (!fs.existsSync(DEBUG_DIR)) {
    fs.mkdirSync(DEBUG_DIR, { recursive: true });
  }
  
  const manager = await createBrowserContext({ headless: false });
  
  try {
    // Authenticate
    console.log('\n1. Authenticating...');
    await ensureAuthenticated(manager.page, manager.context);
    
    // Take initial screenshot
    console.log('\n2. Taking initial screenshot...');
    await manager.page.screenshot({ 
      path: path.join(DEBUG_DIR, '01-initial.png'),
      fullPage: true 
    });
    
    // Wait for page to stabilise
    await manager.page.waitForTimeout(3000);
    
    // Open search with keyboard shortcut
    console.log('\n3. Opening search (Cmd+E)...');
    const isMac = process.platform === 'darwin';
    await manager.page.keyboard.press(isMac ? 'Meta+e' : 'Control+e');
    await manager.page.waitForTimeout(2000);
    
    await manager.page.screenshot({ 
      path: path.join(DEBUG_DIR, '02-search-opened.png'),
      fullPage: true 
    });
    
    // Find and list all inputs
    console.log('\n4. Scanning for input elements...');
    const inputs = await manager.page.locator('input').all();
    console.log(`   Found ${inputs.length} input elements:`);
    
    for (let i = 0; i < inputs.length; i++) {
      const input = inputs[i];
      try {
        const isVisible = await input.isVisible();
        if (!isVisible) continue;
        
        const placeholder = await input.getAttribute('placeholder') ?? '';
        const ariaLabel = await input.getAttribute('aria-label') ?? '';
        const dataTid = await input.getAttribute('data-tid') ?? '';
        const type = await input.getAttribute('type') ?? '';
        const id = await input.getAttribute('id') ?? '';
        
        console.log(`   [${i}] visible=true, placeholder="${placeholder}", aria-label="${ariaLabel}", data-tid="${dataTid}", type="${type}", id="${id}"`);
      } catch {
        // Skip
      }
    }
    
    // Try to find and use search input
    console.log('\n5. Looking for search input...');
    const searchSelectors = [
      'input[placeholder*="Search" i]',
      'input[aria-label*="Search" i]',
      '[data-tid*="search"] input',
      'input[type="search"]',
    ];
    
    let searchInput = null;
    for (const sel of searchSelectors) {
      const loc = manager.page.locator(sel).first();
      if (await loc.count() > 0 && await loc.isVisible()) {
        console.log(`   Found with selector: ${sel}`);
        searchInput = loc;
        break;
      }
    }
    
    if (searchInput) {
      // Type the query
      console.log('\n6. Typing query...');
      await searchInput.fill(query);
      await manager.page.waitForTimeout(500);
      
      await manager.page.screenshot({ 
        path: path.join(DEBUG_DIR, '03-query-typed.png'),
        fullPage: true 
      });
      
      // Submit
      console.log('\n7. Submitting search (Enter)...');
      await manager.page.keyboard.press('Enter');
      await manager.page.waitForTimeout(5000);
      
      await manager.page.screenshot({ 
        path: path.join(DEBUG_DIR, '04-results.png'),
        fullPage: true 
      });
      
      // Scan for result elements
      console.log('\n8. Scanning for result elements...');
      const resultSelectors = [
        '[data-tid*="search"]',
        '[data-tid*="result"]',
        '[role="listitem"]',
        '[role="option"]',
        '.search-result',
        '[data-tid*="message"]',
      ];
      
      for (const sel of resultSelectors) {
        const count = await manager.page.locator(sel).count();
        if (count > 0) {
          console.log(`   ${sel}: ${count} elements`);
          
          // Get first element's text content preview
          const first = manager.page.locator(sel).first();
          const text = await first.textContent().catch(() => null);
          if (text) {
            console.log(`      Preview: "${text.substring(0, 80).replace(/\n/g, ' ')}..."`);
          }
        }
      }
      
      // Dump page HTML structure (simplified)
      console.log('\n9. Dumping main content structure...');
      const mainContent = await manager.page.locator('main, [role="main"], #app, .app-container').first();
      if (await mainContent.count() > 0) {
        const html = await mainContent.innerHTML();
        fs.writeFileSync(
          path.join(DEBUG_DIR, 'page-structure.html'),
          html.substring(0, 50000) // First 50KB
        );
        console.log('   Saved to debug-output/page-structure.html');
      }
      
    } else {
      console.log('   ❌ Could not find search input');
    }
    
    console.log('\n10. Keeping browser open for 30 seconds for manual inspection...');
    console.log(`    Screenshots saved to: ${DEBUG_DIR}`);
    await manager.page.waitForTimeout(30000);
    
  } finally {
    await closeBrowser(manager, true);
  }
  
  console.log('\n✅ Debug session complete');
}

main().catch(console.error);
