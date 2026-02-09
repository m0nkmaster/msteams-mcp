/**
 * Teams Chat Export Script v2.0
 * 
 * Exports Microsoft Teams chat messages to Markdown format.
 * Run this in the browser console while viewing a Teams chat.
 * 
 * Features:
 * - **API-first approach**: Uses Teams' internal chatsvc API for fast, reliable export
 * - Falls back to DOM scraping if API fails
 * - Captures sender, timestamp, content
 * - Generates deep links to each message
 * - Preserves reactions (via DOM fallback)
 * - Detects edited messages
 * - Filters by configurable date range
 * - Detects open threads and offers to export just the thread
 * 
 * Usage:
 * 1. Open Teams in browser (teams.microsoft.com)
 * 2. Navigate to the chat/channel to export
 * 3. Open DevTools (F12) → Console tab
 * 4. Paste this entire script and press Enter
 * 5. Configure days to capture
 * 6. Click Export
 */
(async () => {
  // ─────────────────────────────────────────────────────────────────────────────
  // Configuration
  // ─────────────────────────────────────────────────────────────────────────────
  // Utility Functions
  // ─────────────────────────────────────────────────────────────────────────────
  
  /** Strips HTML tags and decodes entities. */
  const stripHtml = (html) => {
    if (!html) return '';
    return html
      .replace(/<[^>]*>/g, ' ')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")
      .replace(/&apos;/g, "'")
      .replace(/\s+/g, ' ')
      .trim();
  };
  
  /** Builds a deep link to a Teams message. */
  const buildMessageLink = (conversationId, messageId) => {
    if (!conversationId || !messageId) return null;
    return `https://teams.microsoft.com/l/message/${encodeURIComponent(conversationId)}/${messageId}`;
  };
  
  /** Extracts conversation ID from the current URL. */
  const getConversationIdFromUrl = () => {
    const url = new URL(window.location.href);
    let convId = null;
    
    // Method 1: URL search params
    convId = url.searchParams.get('ctx')?.match(/ctx=([^&]+)/)?.[1];
    if (!convId) convId = url.searchParams.get('conversationId');
    if (!convId) convId = url.searchParams.get('threadId');
    if (convId) { console.log('[ConvID] Found via URL params:', convId); return convId; }
    
    // Method 2: Hash fragment
    if (url.hash) {
      const hashMatch = url.hash.match(/conversationId=([^&]+)/);
      if (hashMatch) {
        convId = decodeURIComponent(hashMatch[1]);
        console.log('[ConvID] Found via hash:', convId);
        return convId;
      }
    }
    
    // Method 3: React component data attributes
    const stateEl = document.querySelector('[data-tid="message-pane-node"]');
    const convAttr = stateEl?.getAttribute('data-convid') || 
                     stateEl?.closest('[data-convid]')?.getAttribute('data-convid');
    if (convAttr) {
      console.log('[ConvID] Found via data-convid:', convAttr);
      return convAttr;
    }
    
    // Method 4: Any element with data-convid
    const anyConvEl = document.querySelector('[data-convid]');
    if (anyConvEl) {
      convId = anyConvEl.getAttribute('data-convid');
      console.log('[ConvID] Found via any data-convid:', convId);
      return convId;
    }
    
    // Method 5: localStorage ts.teams
    try {
      const teamsState = localStorage.getItem('ts.teams');
      if (teamsState) {
        const parsed = JSON.parse(teamsState);
        convId = parsed?.selectedConversation?.id;
        if (convId) {
          console.log('[ConvID] Found via ts.teams localStorage:', convId);
          return convId;
        }
      }
    } catch (e) { /* ignore */ }
    
    // Method 6: Look for channel ID in any visible link
    const channelLinks = document.querySelectorAll('a[href*="thread.tacv2"]');
    for (const link of channelLinks) {
      const match = link.href.match(/(19:[^/]+@thread\.tacv2)/);
      if (match) {
        convId = decodeURIComponent(match[1]);
        console.log('[ConvID] Found via channel link:', convId);
        return convId;
      }
    }
    
    // Method 7: Extract from message links in the DOM (already rendered messages)
    const msgLinks = document.querySelectorAll('a[href*="/l/message/"]');
    for (const link of msgLinks) {
      const match = link.href.match(/\/l\/message\/([^/]+)\//);
      if (match) {
        convId = decodeURIComponent(match[1]);
        console.log('[ConvID] Found via message link:', convId);
        return convId;
      }
    }
    
    // Method 8: Look in sessionStorage
    try {
      for (let i = 0; i < sessionStorage.length; i++) {
        const key = sessionStorage.key(i);
        if (key?.includes('conversation') || key?.includes('thread')) {
          const val = sessionStorage.getItem(key);
          const match = val?.match(/(19:[^"]+@thread\.tacv2)/);
          if (match) {
            convId = match[1];
            console.log('[ConvID] Found via sessionStorage:', convId);
            return convId;
          }
        }
      }
    } catch (e) { /* ignore */ }
    
    console.log('[ConvID] Could not detect conversation ID');
    return null;
  };
  
  /** Gets the chat title from the page. */
  const getChatTitle = () => {
    return document.querySelector('h2')?.textContent || 
           document.querySelector('[data-tid="chat-header-title"]')?.textContent ||
           'Teams Chat';
  };
  
  // ─────────────────────────────────────────────────────────────────────────────
  // DOM Scraping Functions
  // ─────────────────────────────────────────────────────────────────────────────
  
  /** Extracts messages from a DOM pane (fallback method). */
  const extractFromPane = (pane, conversationId) => {
    const msgs = [];
    pane.querySelectorAll('[data-tid="chat-pane-item"]').forEach(item => {
      const msg = item.querySelector('[data-tid="chat-pane-message"]');
      const ctrl = item.querySelector('[data-tid="control-message-renderer"]');
      let sender = '', isoDate = '', content = '', edited = false, links = [], reactions = [];
      let messageId = null;
      
      if (ctrl) {
        sender = '[System]';
        content = ctrl.textContent?.trim() || '';
        isoDate = item.querySelector('time')?.getAttribute('datetime') || '';
      } else if (msg) {
        sender = item.querySelector('[data-tid="message-author-name"]')?.textContent?.trim() || '';
        const timeEl = item.querySelector('[id^="timestamp-"]') || item.querySelector('time');
        isoDate = timeEl?.getAttribute('datetime') || '';
        content = msg.querySelector('[id^="content-"]:not([id^="content-control"])')?.textContent?.trim() || '';
        edited = !!item.querySelector('[id^="edited-"]');
        
        // Extract message ID from timestamp (epoch ms)
        if (isoDate) {
          try {
            messageId = String(new Date(isoDate).getTime());
          } catch (e) { /* ignore */ }
        }
        
        links = [...msg.querySelectorAll('a[href]')]
          .map(a => ({ text: a.textContent?.substring(0, 80), url: a.href }))
          .filter(l => l.url && !l.url.includes('statics.teams') && !l.url.startsWith('javascript'));
        reactions = [...msg.querySelectorAll('[data-tid="diverse-reaction-pill-button"]')]
          .map(r => r.textContent?.trim()).filter(Boolean);
      }
      
      if (content) {
        msgs.push({
          id: messageId,
          sender,
          isoDate,
          content,
          edited,
          links: links.length ? links : null,
          reactions: reactions.length ? reactions : null,
          messageLink: messageId && conversationId ? buildMessageLink(conversationId, messageId) : null,
          conversationId,
        });
      }
    });
    return msgs;
  };
  
  /** Scrolls through chat and extracts messages via DOM (fallback). */
  const extractMessagesViaDom = async (conversationId, options = {}) => {
    const { cutoffDate, onProgress, expandThreads = false } = options;
    
    const messages = new Map();
    const chatPane = document.getElementById('chat-pane-list');
    if (!chatPane) {
      throw new Error('No chat pane found. Make sure a chat is open.');
    }
    
    const viewport = chatPane.parentElement;
    
    // Extract messages from main chat pane
    const extract = () => {
      chatPane.querySelectorAll('[data-tid="chat-pane-item"]').forEach(item => {
        const msg = item.querySelector('[data-tid="chat-pane-message"]');
        const ctrl = item.querySelector('[data-tid="control-message-renderer"]');
        let sender = '', timeDisplay = '', isoDate = '', content = '';
        let edited = false, links = [], reactions = [], threadInfo = null;
        let isThreadPreview = false, isSentToChannel = false;
        let messageId = null;
        
        if (ctrl) {
          sender = '[System]';
          content = ctrl.textContent?.trim() || '';
          isoDate = item.querySelector('time')?.getAttribute('datetime') || '';
        } else if (msg) {
          sender = item.querySelector('[data-tid="message-author-name"]')?.textContent?.trim() || '';
          const timeEl = item.querySelector('[id^="timestamp-"]') || item.querySelector('time');
          timeDisplay = timeEl?.textContent?.trim() || '';
          isoDate = timeEl?.getAttribute('datetime') || '';
          content = msg.querySelector('[id^="content-"]:not([id^="content-control"])')?.textContent?.trim() || '';
          
          // Extract message ID
          if (isoDate) {
            try {
              messageId = String(new Date(isoDate).getTime());
            } catch (e) { /* ignore */ }
          }
          
          // Track "Also sent to channel" messages
          if (content.match(/^Also sent to channel\s*/i)) {
            isSentToChannel = true;
            content = content.replace(/^Also sent to channel\s*/i, '');
          }
          edited = !!item.querySelector('[id^="edited-"]');
          
          links = [...msg.querySelectorAll('a[href]')]
            .map(a => ({ text: a.textContent?.substring(0, 80), url: a.href }))
            .filter(l => l.url && !l.url.includes('statics.teams') && !l.url.startsWith('javascript'));
          
          reactions = [...msg.querySelectorAll('[data-tid="diverse-reaction-pill-button"]')]
            .map(r => r.textContent?.trim()).filter(Boolean);
          
          // Check if this is a "Replied in thread" preview
          if (content.startsWith('Replied in thread:')) {
            isThreadPreview = true;
            content = content.replace(/^Replied in thread:\s*/, '');
          }
          
          // Check for thread replies
          const replySummary = item.querySelector('[data-tid="replies-summary-authors"]');
          if (replySummary && !isThreadPreview) {
            const summaryParent = replySummary.closest('[class*="repl"]') || replySummary.parentElement?.parentElement;
            const summaryText = summaryParent?.textContent || '';
            
            const replyMatch = summaryText.match(/\b(\d+)\s*repl(?:y|ies)/i);
            if (replyMatch) {
              const lastReplyMatch = summaryText.match(/Last reply\s+([^F]+?)(?:Follow|$)/i);
              threadInfo = {
                replyCount: parseInt(replyMatch[1]),
                lastReply: lastReplyMatch ? lastReplyMatch[1].trim() : null,
                replies: []
              };
            }
          }
        }
        
        if (content) {
          const key = `${sender}-${isoDate || timeDisplay}-${content.substring(0, 40)}`;
          if (!messages.has(key)) {
            messages.set(key, {
              id: messageId,
              sender,
              timeDisplay,
              isoDate,
              content,
              edited,
              links: links.length ? links : null,
              reactions: reactions.length ? reactions : null,
              threadInfo,
              isThreadPreview,
              isSentToChannel,
              contentSnippet: content.substring(0, 50),
              messageLink: messageId && conversationId ? buildMessageLink(conversationId, messageId) : null,
              conversationId,
            });
          }
        }
      });
    };

    // Scroll through chat and collect messages
    viewport.scrollTop = viewport.scrollHeight;
    await new Promise(r => setTimeout(r, 500));
    
    let pos = viewport.scrollHeight;
    let done = false;
    
    while (pos > 0 && !done) {
      viewport.scrollTop = pos;
      await new Promise(r => setTimeout(r, 150));
      extract();
      
      if (onProgress) {
        onProgress(`Scanning via DOM... ${messages.size} messages`);
      }
      
      for (const m of messages.values()) {
        if (m.isoDate && cutoffDate && new Date(m.isoDate) < cutoffDate) {
          done = true;
          break;
        }
      }
      pos -= 300;
    }
    extract();
    
    // Filter and return
    const result = [...messages.values()]
      .filter(m => !m.isThreadPreview)
      .filter(m => !expandThreads || !m.isSentToChannel)
      .filter(m => {
        if (m.isoDate && cutoffDate) return new Date(m.isoDate) >= cutoffDate;
        return true;
      });
    
    console.log(`[DOM] Extracted ${result.length} messages`);
    return { messages: result, source: 'dom' };
  };
  
  // ─────────────────────────────────────────────────────────────────────────────
  // Markdown Generation
  // ─────────────────────────────────────────────────────────────────────────────
  
  /** Generates markdown from messages array. */
  const generateMarkdown = (messages, title, options = {}) => {
    const { isThread = false, includeLinks = true } = options;
    
    // Sort chronologically
    messages.sort((a, b) => {
      const da = a.isoDate ? new Date(a.isoDate).getTime() : 0;
      const db = b.isoDate ? new Date(b.isoDate).getTime() : 0;
      return da - db;
    });

    let md = '# ' + title + (isThread ? ' (Thread)' : '') + '\n\n';
    md += '**Exported:** ' + new Date().toLocaleDateString('en-GB') + '\n';
    md += '**Messages:** ' + messages.length + '\n\n---\n\n';
    
    let lastDate = '';
    
    for (const m of messages) {
      let dateStr = '', timeStr = '';
      
      if (m.isoDate) {
        const d = new Date(m.isoDate);
        dateStr = d.toLocaleDateString('en-GB', {
          weekday: 'long', day: 'numeric', month: 'long', year: 'numeric'
        });
        timeStr = d.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
      } else {
        dateStr = 'Unknown Date';
        timeStr = '';
      }
      
      // Date header
      if (dateStr !== lastDate) {
        if (lastDate) md += '\n---\n\n';
        md += '## ' + dateStr + '\n\n';
        lastDate = dateStr;
      }
      
      // Message header with optional link
      md += '**' + m.sender + '**';
      if (timeStr) md += ' (' + timeStr + ')';
      if (m.edited) md += ' *(edited)*';
      if (includeLinks && m.messageLink) {
        md += ' [[link](' + m.messageLink + ')]';
      }
      md += ':\n';
      
      // Content
      md += m.content.split('\n').map(l => '> ' + l).join('\n') + '\n';
      
      // Links (from DOM scraping)
      if (m.links) {
        md += '>\n> 🔗 ' + m.links.map(l => '[' + l.text + '](' + l.url + ')').join(' | ') + '\n';
      }
      
      // Reactions (from DOM scraping)
      if (m.reactions) {
        md += '>\n> ' + m.reactions.join(' | ') + '\n';
      }
      
      // Thread replies (for full chat export)
      if (m.threadInfo) {
        if (m.threadInfo.replies && m.threadInfo.replies.length > 0) {
          md += '>\n> 💬 **Thread (' + m.threadInfo.replies.length + ' replies):**\n';
          for (const reply of m.threadInfo.replies) {
            let replyTime = '';
            if (reply.isoDate) {
              replyTime = new Date(reply.isoDate).toLocaleTimeString('en-GB', {
                hour: '2-digit', minute: '2-digit'
              });
            }
            md += '>\n> > **' + reply.sender + '**';
            if (replyTime) md += ' (' + replyTime + ')';
            md += ':\n';
            md += reply.content.split('\n').map(l => '> > > ' + l).join('\n') + '\n';
            if (reply.reactions) {
              md += '> > >\n> > > ' + reply.reactions.join(' | ') + '\n';
            }
          }
        } else {
          md += '>\n> 💬 **' + m.threadInfo.replyCount + ' ';
          md += (m.threadInfo.replyCount === 1 ? 'reply' : 'replies') + '**';
          if (m.threadInfo.lastReply) md += ' *(last: ' + m.threadInfo.lastReply + ')*';
          md += ' *[not expanded]*\n';
        }
      }
      
      md += '\n';
    }
    
    return md;
  };
  
  // ─────────────────────────────────────────────────────────────────────────────
  // UI Creation
  // ─────────────────────────────────────────────────────────────────────────────
  
  const chatTitle = getChatTitle();
  const conversationId = getConversationIdFromUrl();
  
  // Check if a thread is already open
  const rightRail = document.querySelector('[data-tid="right-rail-message-pane-body"]');
  const threadIsOpen = rightRail && rightRail.offsetParent !== null;
  
  // Create overlay
  const overlay = document.createElement('div');
  Object.assign(overlay.style, {
    position: 'fixed', top: '0', left: '0', right: '0', bottom: '0',
    background: 'rgba(0,0,0,0.5)', zIndex: '999998'
  });
  
  // Create modal
  const modal = document.createElement('div');
  Object.assign(modal.style, {
    position: 'fixed', top: '50%', left: '50%', transform: 'translate(-50%,-50%)',
    background: '#fff', padding: '24px', borderRadius: '12px',
    boxShadow: '0 8px 32px rgba(0,0,0,0.3)', zIndex: '999999',
    fontFamily: 'system-ui', minWidth: '450px', maxWidth: '90vw'
  });
  
  const close = () => { overlay.remove(); modal.remove(); };
  
  // Helper to scroll through a thread pane and extract all messages
  const extractAllFromThread = async (pane) => {
    const msgs = new Map();
    const extract = () => {
      pane.querySelectorAll('[data-tid="chat-pane-item"]').forEach(item => {
        const msg = item.querySelector('[data-tid="chat-pane-message"]');
        if (msg) {
          const sender = item.querySelector('[data-tid="message-author-name"]')?.textContent?.trim() || '';
          const timeEl = item.querySelector('[id^="timestamp-"]') || item.querySelector('time');
          const isoDate = timeEl?.getAttribute('datetime') || '';
          const content = msg.querySelector('[id^="content-"]:not([id^="content-control"])')?.textContent?.trim() || '';
          const edited = !!item.querySelector('[id^="edited-"]');
          
          let messageId = null;
          if (isoDate) {
            try {
              messageId = String(new Date(isoDate).getTime());
            } catch (e) { /* ignore */ }
          }
          
          const links = [...msg.querySelectorAll('a[href]')]
            .map(a => ({ text: a.textContent?.substring(0, 80), url: a.href }))
            .filter(l => l.url && !l.url.includes('statics.teams') && !l.url.startsWith('javascript'));
          const reactions = [...msg.querySelectorAll('[data-tid="diverse-reaction-pill-button"]')]
            .map(r => r.textContent?.trim()).filter(Boolean);
          
          if (content) {
            const key = `${sender}-${isoDate}-${content.substring(0, 40)}`;
            if (!msgs.has(key)) {
              msgs.set(key, {
                id: messageId,
                sender,
                isoDate,
                content,
                edited,
                links: links.length ? links : null,
                reactions: reactions.length ? reactions : null,
                messageLink: messageId && conversationId ? buildMessageLink(conversationId, messageId) : null,
              });
            }
          }
        }
      });
    };
    
    // Find the scrollable container
    let scrollContainer = pane;
    let el = pane;
    while (el && el !== document.body) {
      if (el.scrollHeight > el.clientHeight + 10) {
        scrollContainer = el;
        break;
      }
      el = el.parentElement;
    }
    
    // Scroll to bottom first, then up
    scrollContainer.scrollTop = scrollContainer.scrollHeight;
    await new Promise(r => setTimeout(r, 300));
    extract();
    
    // Scroll up through thread to load older messages
    let scrollAttempts = 0;
    const maxScrollAttempts = 50;
    while (scrollAttempts < maxScrollAttempts) {
      const prevCount = msgs.size;
      const prevScroll = scrollContainer.scrollTop;
      
      scrollContainer.scrollTop -= 400;
      await new Promise(r => setTimeout(r, 200));
      extract();
      
      if (msgs.size === prevCount && (scrollContainer.scrollTop === prevScroll || scrollContainer.scrollTop === 0)) {
        break;
      }
      scrollAttempts++;
    }
    
    return [...msgs.values()].sort((a, b) => {
      const da = a.isoDate ? new Date(a.isoDate).getTime() : 0;
      const db = b.isoDate ? new Date(b.isoDate).getTime() : 0;
      return da - db;
    });
  };
  
  /**
   * Expands threads via DOM interaction - clicks each thread, extracts replies, closes it.
   * This works around auth issues with the API by using the actual Teams UI.
   */
  const expandThreadsViaDom = async (messages, conversationId, onProgress) => {
    const chatPane = document.getElementById('chat-pane-list');
    if (!chatPane) {
      console.log('[DOM Threads] No chat pane found');
      return;
    }
    
    let threadsExpanded = 0;
    let totalReplies = 0;
    
    // Find messages that have thread info (detected during DOM scrape)
    const messagesWithThreads = messages.filter(m => m.threadInfo && m.threadInfo.replyCount > 0);
    console.log(`[DOM Threads] Found ${messagesWithThreads.length} messages with threads`);
    
    for (let i = 0; i < messagesWithThreads.length; i++) {
      const msg = messagesWithThreads[i];
      if (onProgress) {
        onProgress(`Expanding thread ${i + 1}/${messagesWithThreads.length}...`);
      }
      
      // Find the message element in the DOM by matching content snippet
      const msgElements = chatPane.querySelectorAll('[data-tid="chat-pane-item"]');
      let targetItem = null;
      
      for (const item of msgElements) {
        const contentEl = item.querySelector('[id^="content-"]:not([id^="content-control"])');
        const content = contentEl?.textContent?.trim() || '';
        // Match by content snippet (first 40 chars)
        if (content && msg.content && content.substring(0, 40) === msg.content.substring(0, 40)) {
          targetItem = item;
          break;
        }
      }
      
      if (!targetItem) {
        console.log(`[DOM Threads] Could not find message: ${msg.content?.substring(0, 30)}...`);
        continue;
      }
      
      // Find the "X replies" button/link to click
      const repliesLink = targetItem.querySelector('[data-tid="replies-summary-authors"]');
      const repliesButton = repliesLink?.closest('button') || 
                            repliesLink?.closest('[role="button"]') ||
                            targetItem.querySelector('[class*="repl"]')?.querySelector('button');
      
      if (!repliesButton) {
        console.log(`[DOM Threads] No replies button found for: ${msg.content?.substring(0, 30)}...`);
        continue;
      }
      
      // Click to open thread
      console.log(`[DOM Threads] Clicking to open thread: ${msg.content?.substring(0, 30)}...`);
      repliesButton.click();
      
      // Wait for thread pane to appear
      await new Promise(r => setTimeout(r, 800));
      
      // Find the thread pane (right rail)
      const threadPane = document.querySelector('[data-tid="right-rail-message-pane-body"]');
      if (!threadPane || threadPane.offsetParent === null) {
        console.log(`[DOM Threads] Thread pane didn't open`);
        continue;
      }
      
      // Extract all messages from the thread
      const threadMessages = await extractAllFromThread(threadPane);
      
      // Filter out the root message (we already have it) and add as replies
      const replies = threadMessages.filter(tm => {
        // Skip if it matches the root message
        if (tm.content?.substring(0, 40) === msg.content?.substring(0, 40)) {
          return false;
        }
        return true;
      });
      
      if (replies.length > 0) {
        msg.threadInfo.replies = replies;
        totalReplies += replies.length;
        console.log(`[DOM Threads] Extracted ${replies.length} replies`);
      }
      
      // Close the thread pane - try multiple selectors
      let closed = false;
      const closeSelectors = [
        '[data-tid="thread-list-pane-toggle-button"]',
        '[data-tid="right-rail-close-button"]',
        '[data-tid="message-pane-close-button"]',
        'button[aria-label="Close"]',
        'button[aria-label="Close thread"]',
        '[class*="closeButton"]',
        '[class*="CloseButton"]',
      ];
      
      for (const selector of closeSelectors) {
        const btn = document.querySelector(selector);
        if (btn && btn.offsetParent !== null) {
          console.log(`[DOM Threads] Closing via: ${selector}`);
          btn.click();
          closed = true;
          break;
        }
      }
      
      if (!closed) {
        // Fallback: try pressing Escape
        console.log('[DOM Threads] No close button found, trying Escape key');
        document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
      }
      
      await new Promise(r => setTimeout(r, 600));
      
      // Verify thread pane closed
      const stillOpen = document.querySelector('[data-tid="right-rail-message-pane-body"]');
      if (stillOpen && stillOpen.offsetParent !== null) {
        console.log('[DOM Threads] Thread pane still open, waiting longer...');
        await new Promise(r => setTimeout(r, 500));
      }
      
      threadsExpanded++;
      
      // Small delay between threads
      await new Promise(r => setTimeout(r, 200));
    }
    
    console.log(`[DOM Threads] Done - expanded ${threadsExpanded} threads, ${totalReplies} replies`);
  };
  
  // If thread is open, show thread export option
  if (threadIsOpen) {
    const visibleCount = rightRail.querySelectorAll('[data-tid="chat-pane-item"]').length;

    const title = document.createElement('h2');
    title.textContent = 'Thread Detected';
    Object.assign(title.style, { margin: '0 0 16px', color: '#242424' });
    
    const info = document.createElement('div');
    Object.assign(info.style, {
      background: '#fff3cd', padding: '12px', borderRadius: '6px', marginBottom: '16px',
      border: '1px solid #ffc107', color: '#856404'
    });
    info.textContent = `A thread is currently open (${visibleCount}+ messages visible).`;
    
    const question = document.createElement('p');
    question.textContent = 'What would you like to export?';
    Object.assign(question.style, { margin: '16px 0', color: '#242424' });
    
    const buttonArea = document.createElement('div');
    Object.assign(buttonArea.style, {
      display: 'flex', flexDirection: 'column', gap: '12px', marginTop: '20px'
    });
    
    const threadBtn = document.createElement('button');
    threadBtn.textContent = '💬 Export This Thread Only';
    Object.assign(threadBtn.style, {
      padding: '12px 20px', border: 'none', borderRadius: '6px',
      cursor: 'pointer', background: '#6264a7', color: '#fff', fontSize: '14px'
    });
    
    const fullChatBtn = document.createElement('button');
    fullChatBtn.textContent = '📋 Export Full Chat (close thread first)';
    Object.assign(fullChatBtn.style, {
      padding: '12px 20px', border: 'none', borderRadius: '6px',
      cursor: 'pointer', background: '#f0f0f0', fontSize: '14px', color: '#242424'
    });
    
    const cancelBtn = document.createElement('button');
    cancelBtn.textContent = 'Cancel';
    Object.assign(cancelBtn.style, {
      padding: '10px 20px', border: 'none', borderRadius: '6px',
      cursor: 'pointer', background: 'transparent', color: '#666', fontSize: '14px'
    });
    
    buttonArea.appendChild(threadBtn);
    buttonArea.appendChild(fullChatBtn);
    buttonArea.appendChild(cancelBtn);
    
    modal.appendChild(title);
    modal.appendChild(info);
    modal.appendChild(question);
    modal.appendChild(buttonArea);
    
    document.body.appendChild(overlay);
    document.body.appendChild(modal);
    
    overlay.onclick = close;
    cancelBtn.onclick = close;
    
    threadBtn.onclick = async () => {
      buttonArea.style.display = 'none';
      question.style.display = 'none';
      info.style.background = '#e3f2fd';
      info.style.border = '1px solid #2196f3';
      info.textContent = 'Scrolling through thread to capture all messages...';
      
      const threadMessages = await extractAllFromThread(rightRail);
      
      const md = generateMarkdown(threadMessages, chatTitle, { isThread: true });
      try {
        await navigator.clipboard.writeText(md);
        info.style.background = '#d4edda';
        info.style.border = '1px solid #28a745';
        info.textContent = `✓ Copied ${threadMessages.length} thread messages to clipboard!`;
        setTimeout(close, 2000);
      } catch (e) {
        console.log('=== MARKDOWN OUTPUT ===\n', md);
        info.textContent = 'Clipboard blocked - check console (F12)';
      }
    };
    
    fullChatBtn.onclick = async () => {
      const toggleBtn = document.querySelector('[data-tid="thread-list-pane-toggle-button"]');
      if (toggleBtn) {
        toggleBtn.click();
        await new Promise(r => setTimeout(r, 500));
      }
      modal.remove();
      showFullChatUI();
    };
    
  } else {
    document.body.appendChild(overlay);
    document.body.appendChild(modal);
    showFullChatUI();
  }
  
  function showFullChatUI() {
    while (modal.firstChild) {
      modal.removeChild(modal.firstChild);
    }
    
    const title = document.createElement('h2');
    title.textContent = 'Export Teams Chat';
    Object.assign(title.style, { margin: '0 0 16px', color: '#242424' });
    
    const info = document.createElement('div');
    Object.assign(info.style, {
      background: '#f5f5f5', padding: '12px', borderRadius: '6px', marginBottom: '16px',
      color: '#242424', fontSize: '13px'
    });
    
    const chatInfo = document.createElement('div');
    const chatLabel = document.createElement('strong');
    chatLabel.textContent = 'Chat: ';
    chatInfo.appendChild(chatLabel);
    chatInfo.appendChild(document.createTextNode(chatTitle));
    info.appendChild(chatInfo);
    
    if (conversationId) {
      const convInfo = document.createElement('div');
      const convLabel = document.createElement('strong');
      convLabel.textContent = 'Conversation ID: ';
      convInfo.appendChild(convLabel);
      const convCode = document.createElement('code');
      Object.assign(convCode.style, { fontSize: '11px', background: '#e0e0e0', padding: '2px 4px', borderRadius: '3px' });
      convCode.textContent = conversationId.substring(0, 40) + (conversationId.length > 40 ? '...' : '');
      convInfo.appendChild(convCode);
      convInfo.style.marginTop = '8px';
      info.appendChild(convInfo);
      
      const methodInfo = document.createElement('div');
      const methodLabel = document.createElement('strong');
      methodLabel.textContent = 'Method: ';
      methodInfo.appendChild(methodLabel);
      const methodSpan = document.createElement('span');
      methodSpan.style.color = '#0078d4';
      methodSpan.textContent = 'DOM scraping';
      methodInfo.appendChild(methodSpan);
      methodInfo.appendChild(document.createTextNode(' with thread expansion'));
      methodInfo.style.marginTop = '8px';
      info.appendChild(methodInfo);
    } else {
      const methodInfo = document.createElement('div');
      const methodLabel = document.createElement('strong');
      methodLabel.textContent = 'Method: ';
      methodInfo.appendChild(methodLabel);
      const methodSpan = document.createElement('span');
      methodSpan.style.color = '#856404';
      methodSpan.textContent = 'DOM scraping';
      methodInfo.appendChild(methodSpan);
      methodInfo.appendChild(document.createTextNode(' (could not detect conversation ID)'));
      methodInfo.style.marginTop = '8px';
      info.appendChild(methodInfo);
    }
    
    const label = document.createElement('label');
    label.textContent = 'Days to capture: ';
    Object.assign(label.style, { display: 'block', marginBottom: '8px', color: '#242424' });
    
    const input = document.createElement('input');
    input.type = 'number';
    input.value = '7';
    input.min = '1';
    input.max = '365';
    Object.assign(input.style, {
      width: '100%', padding: '8px', border: '1px solid #ddd',
      borderRadius: '6px', marginTop: '4px', boxSizing: 'border-box',
      color: '#242424', background: '#fff'
    });
    label.appendChild(input);
    
    // Include links checkbox
    const linksLabel = document.createElement('label');
    Object.assign(linksLabel.style, {
      display: 'flex', alignItems: 'center', gap: '8px', marginTop: '12px', cursor: 'pointer',
      color: '#242424'
    });
    const linksCheck = document.createElement('input');
    linksCheck.type = 'checkbox';
    linksCheck.checked = true;
    linksLabel.appendChild(linksCheck);
    linksLabel.appendChild(document.createTextNode('Include message links (click to open in Teams)'));
    
    // Expand threads checkbox (for channels)
    const threadsLabel = document.createElement('label');
    Object.assign(threadsLabel.style, {
      display: 'flex', alignItems: 'center', gap: '8px', marginTop: '8px', cursor: 'pointer',
      color: '#242424'
    });
    const threadsCheck = document.createElement('input');
    threadsCheck.type = 'checkbox';
    threadsCheck.checked = true;
    threadsLabel.appendChild(threadsCheck);
    threadsLabel.appendChild(document.createTextNode('Include thread replies (for channels)'));
    
    const progressArea = document.createElement('div');
    progressArea.style.display = 'none';
    progressArea.style.marginTop = '16px';
    
    const progressBarOuter = document.createElement('div');
    Object.assign(progressBarOuter.style, {
      height: '8px', background: '#e0e0e0', borderRadius: '4px', overflow: 'hidden'
    });
    
    const progressBar = document.createElement('div');
    Object.assign(progressBar.style, {
      height: '100%', background: '#6264a7', width: '0%', transition: 'width 0.3s'
    });
    progressBarOuter.appendChild(progressBar);
    
    const progressText = document.createElement('div');
    Object.assign(progressText.style, { fontSize: '12px', color: '#666', marginTop: '8px' });
    progressText.textContent = 'Starting...';
    
    progressArea.appendChild(progressBarOuter);
    progressArea.appendChild(progressText);
    
    const buttonArea = document.createElement('div');
    Object.assign(buttonArea.style, {
      display: 'flex', gap: '12px', justifyContent: 'flex-end', marginTop: '20px'
    });
    
    const cancelBtn = document.createElement('button');
    cancelBtn.textContent = 'Cancel';
    Object.assign(cancelBtn.style, {
      padding: '10px 20px', border: 'none', borderRadius: '6px',
      cursor: 'pointer', background: '#f0f0f0', color: '#242424'
    });
    
    const exportBtn = document.createElement('button');
    exportBtn.textContent = 'Export';
    Object.assign(exportBtn.style, {
      padding: '10px 20px', border: 'none', borderRadius: '6px',
      cursor: 'pointer', background: '#6264a7', color: '#fff'
    });
    
    buttonArea.appendChild(cancelBtn);
    buttonArea.appendChild(exportBtn);
    
    modal.appendChild(title);
    modal.appendChild(info);
    modal.appendChild(label);
    modal.appendChild(linksLabel);
    modal.appendChild(threadsLabel);
    modal.appendChild(progressArea);
    modal.appendChild(buttonArea);
    
    if (!document.body.contains(modal)) {
      document.body.appendChild(modal);
    }
    
    cancelBtn.onclick = close;
    overlay.onclick = close;

    exportBtn.onclick = async () => {
      const days = parseInt(input.value) || 7;
      const includeLinks = linksCheck.checked;
      const expandThreads = threadsCheck.checked;
      const cutoffDate = new Date();
      cutoffDate.setDate(cutoffDate.getDate() - days);
      cutoffDate.setHours(0, 0, 0, 0);
      
      buttonArea.style.display = 'none';
      progressArea.style.display = 'block';
      
      let result;
      
      // Use DOM scraping directly (API requires auth tokens that aren't available from console)
      progressText.textContent = 'Scanning messages...';
      progressBar.style.width = '20%';
      
      try {
        result = await extractMessagesViaDom(conversationId, {
          cutoffDate,
          onProgress: (msg) => {
            progressText.textContent = msg;
          }
        });
      } catch (e) {
        progressText.textContent = 'Error: ' + e.message;
        progressText.style.color = 'red';
        return;
      }
      
      // Expand thread replies if enabled and this is a channel
      const isChannel = conversationId && conversationId.includes('@thread.tacv2');
      console.log(`[Export] expandThreads=${expandThreads}, isChannel=${isChannel}, conversationId=${conversationId}`);
      
      // Check if any messages have threads that need expanding
      const messagesWithThreads = result.messages.filter(m => m.threadInfo && m.threadInfo.replyCount > 0 && (!m.threadInfo.replies || m.threadInfo.replies.length === 0));
      
      if (expandThreads && isChannel && messagesWithThreads.length > 0) {
        progressBar.style.width = '50%';
        progressText.textContent = `Expanding ${messagesWithThreads.length} threads via UI...`;
        
        console.log(`[Threads] Expanding ${messagesWithThreads.length} threads via DOM`);
        
        // Use DOM-based thread expansion (clicks into each thread)
        await expandThreadsViaDom(result.messages, conversationId, (msg) => {
          progressText.textContent = msg;
        });
      }
      
      progressBar.style.width = '90%';
      progressText.textContent = `Generating markdown (${result.messages.length} messages via ${result.source})...`;

      const md = generateMarkdown(result.messages, chatTitle, { includeLinks });

      // Count total including replies
      let totalReplies = 0;
      for (const m of result.messages) {
        if (m.threadInfo?.replies) {
          totalReplies += m.threadInfo.replies.length;
        }
      }
      const totalCount = result.messages.length + totalReplies;
      
      // Copy to clipboard
      try {
        await navigator.clipboard.writeText(md);
        progressBar.style.width = '100%';
        progressText.textContent = `✓ Copied ${result.messages.length} messages + ${totalReplies} replies to clipboard!`;
        progressText.style.color = 'green';
        progressText.style.fontWeight = 'bold';
        setTimeout(close, 2500);
      } catch (e) {
        console.log('=== MARKDOWN OUTPUT ===\n', md);
        progressBar.style.width = '100%';
        progressText.textContent = `✓ ${result.messages.length} messages + ${totalReplies} replies ready - check console (F12)`;
        progressText.style.color = '#856404';
        progressText.style.fontWeight = 'bold';
        
        // Add a "Close" button since we can't auto-close
        const closeBtn = document.createElement('button');
        closeBtn.textContent = 'Close';
        Object.assign(closeBtn.style, {
          marginTop: '16px', padding: '10px 20px', border: 'none', borderRadius: '6px',
          cursor: 'pointer', background: '#6264a7', color: '#fff'
        });
        closeBtn.onclick = close;
        progressArea.appendChild(closeBtn);
      }
    };
  }
})();
