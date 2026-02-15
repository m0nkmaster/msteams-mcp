/**
 * Markdown to Teams HTML conversion utilities.
 */

import { escapeHtmlChars } from './parsers-html.js';

/**
 * Converts inline markdown formatting to Teams HTML within a single line.
 * Handles: bold, italic, strikethrough, inline code.
 * Text outside of formatting markers is HTML-escaped.
 */
function convertInlineFormatting(line: string): string {
  // Process inline code first (to prevent other formatting inside code spans)
  // Split on `code` patterns, escape and format alternately
  const codeParts = line.split(/`([^`]+)`/);
  let result = '';
  
  for (let i = 0; i < codeParts.length; i++) {
    if (i % 2 === 1) {
      // Inside backticks - render as code, only escape HTML
      result += `<code>${escapeHtmlChars(codeParts[i])}</code>`;
    } else {
      // Outside backticks - process other inline formatting
      let segment = escapeHtmlChars(codeParts[i]);
      
      // Bold: **text** or __text__
      segment = segment.replace(/\*\*(.+?)\*\*/g, '<b>$1</b>');
      segment = segment.replace(/__(.+?)__/g, '<b>$1</b>');
      
      // Italic: *text* or _text_ (but not inside words for underscore)
      segment = segment.replace(/\*(.+?)\*/g, '<i>$1</i>');
      segment = segment.replace(/(?<!\w)_(.+?)_(?!\w)/g, '<i>$1</i>');
      
      // Strikethrough: ~~text~~
      segment = segment.replace(/~~(.+?)~~/g, '<s>$1</s>');
      
      result += segment;
    }
  }
  
  return result;
}

/**
 * Converts markdown-formatted text to Teams-compatible HTML.
 * 
 * Supports:
 * - **bold** / __bold__ → <b>
 * - *italic* / _italic_ → <i>
 * - ~~strikethrough~~ → <s>
 * - `inline code` → <code>
 * - ```code blocks``` → <pre><code>
 * - Newlines → paragraph breaks
 * - Ordered lists (1. item) → <ol><li>
 * - Unordered lists (- item, * item) → <ul><li>
 * 
 * Plain text without any formatting is returned as-is (HTML-escaped).
 */
export function markdownToTeamsHtml(text: string): string {
  // Handle fenced code blocks first (```...```)
  // Split text into code blocks and non-code-block segments
  const segments: { type: 'text' | 'codeblock'; content: string; lang?: string }[] = [];
  const codeBlockRegex = /```(\w*)\n?([\s\S]*?)```/g;
  let lastIndex = 0;
  let match;
  
  while ((match = codeBlockRegex.exec(text)) !== null) {
    // Text before this code block
    if (match.index > lastIndex) {
      segments.push({ type: 'text', content: text.substring(lastIndex, match.index) });
    }
    segments.push({ type: 'codeblock', content: match[2], lang: match[1] || undefined });
    lastIndex = match.index + match[0].length;
  }
  // Remaining text after last code block
  if (lastIndex < text.length) {
    segments.push({ type: 'text', content: text.substring(lastIndex) });
  }
  
  const htmlParts: string[] = [];
  
  for (const segment of segments) {
    if (segment.type === 'codeblock') {
      // Code blocks: escape HTML but preserve whitespace
      const escaped = escapeHtmlChars(segment.content.replace(/\n$/, ''));
      htmlParts.push(`<pre><code>${escaped}</code></pre>`);
      continue;
    }
    
    // Process text segments: split into paragraphs on double newlines
    const paragraphs = segment.content.split(/\n{2,}/);
    
    for (const para of paragraphs) {
      const trimmed = para.trim();
      if (!trimmed) continue;
      
      const lines = trimmed.split('\n');
      
      // Check if this paragraph is a list
      const isUnorderedList = lines.every(l => /^\s*[-*]\s+/.test(l));
      const isOrderedList = lines.every(l => /^\s*\d+[.)]\s+/.test(l));
      
      if (isUnorderedList) {
        const items = lines.map(l => {
          const content = l.replace(/^\s*[-*]\s+/, '');
          return `<li>${convertInlineFormatting(content)}</li>`;
        });
        htmlParts.push(`<ul>${items.join('')}</ul>`);
      } else if (isOrderedList) {
        const items = lines.map(l => {
          const content = l.replace(/^\s*\d+[.)]\s+/, '');
          return `<li>${convertInlineFormatting(content)}</li>`;
        });
        htmlParts.push(`<ol>${items.join('')}</ol>`);
      } else {
        // Regular paragraph - join lines with <br>
        const htmlLines = lines.map(l => convertInlineFormatting(l));
        htmlParts.push(`<p>${htmlLines.join('<br>')}</p>`);
      }
    }
  }
  
  return htmlParts.join('') || '<p></p>';
}

/**
 * Checks whether text contains any markdown formatting that would
 * benefit from conversion to HTML.
 */
export function hasMarkdownFormatting(text: string): boolean {
  // Code blocks
  if (/```[\s\S]*```/.test(text)) return true;
  // Inline code
  if (/`[^`]+`/.test(text)) return true;
  // Bold
  if (/\*\*.+?\*\*/.test(text) || /__.+?__/.test(text)) return true;
  // Italic (single * or _)
  if (/(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)/.test(text)) return true;
  // Strikethrough
  if (/~~.+?~~/.test(text)) return true;
  // Lists
  if (/^\s*[-*]\s+/m.test(text)) return true;
  if (/^\s*\d+[.)]\s+/m.test(text)) return true;
  // Multiple newlines (paragraph breaks)
  if (/\n/.test(text)) return true;
  
  return false;
}
