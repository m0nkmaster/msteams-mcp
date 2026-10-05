/**
 * Markdown to Teams HTML conversion utilities.
 *
 * Uses `marked` for block/inline parsing (GFM: tables, nested lists, task
 * lists, blockquotes, headings, etc.) with a renderer that emits the shapes
 * the Teams client accepts:
 *  - Tables → bare <table><tbody><tr><td><p>…</p></td>… (no <thead>/<th>,
 *    every cell wrapped in <p> — verified against raw chatsvc content of a
 *    manually composed table message).
 *  - No alignment attributes (Teams ignores them).
 *  - **bold** → <b>, *italic* → <i> (not <strong>/<em> — Teams renders these).
 *  - All raw HTML is stripped/escaped: model-authored content must never be
 *    injected as markup (marked v18 passes HTML through even with html:false,
 *    and emits javascript: hrefs, so sanitizing here is load-bearing).
 *  - Code blocks → <pre><code> without language class (Teams has no syntax
 *    highlighter; the language tag is dropped).
 */

import { Marked, type Tokens } from 'marked';
import { escapeHtmlChars, sanitizeLinkUrl } from './parsers-html.js';

const md = new Marked({ gfm: true, breaks: true });

/** Renders inline tokens to Teams HTML with full escaping. */
function renderInline(tokens: Tokens.Generic[] | undefined): string {
  if (!tokens) return '';
  let out = '';
  for (const tok of tokens) {
    switch (tok.type) {
      case 'strong':
        out += `<b>${renderInline(tok.tokens)}</b>`;
        break;
      case 'em':
        out += `<i>${renderInline(tok.tokens)}</i>`;
        break;
      case 'del':
        out += `<s>${renderInline(tok.tokens)}</s>`;
        break;
      case 'codespan':
        out += `<code>${escapeHtmlChars(tok.text)}</code>`;
        break;
      case 'link': {
        const safeUrl = sanitizeLinkUrl(tok.href).replace(/"/g, '&quot;');
        out += `<a href="${safeUrl}">${renderInline(tok.tokens)}</a>`;
        break;
      }
      case 'image':
        // Teams chat doesn't render agent-supplied <img>; keep alt text.
        out += escapeHtmlChars(tok.text ?? '');
        break;
      case 'br':
        out += '<br>';
        break;
      case 'escape':
        out += escapeHtmlChars(tok.text);
        break;
      default:
        out += escapeHtmlChars(tok.raw ?? '');
    }
  }
  return out;
}

/** Renders block tokens (paragraphs, lists, tables, code, quotes, headings). */
function renderBlocks(tokens: Tokens.Generic[]): string {
  let out = '';
  for (const tok of tokens) {
    switch (tok.type) {
      case 'paragraph':
        out += `<p>${renderInline(tok.tokens)}</p>`;
        break;
      case 'heading':
        out += `<p><b>${renderInline(tok.tokens)}</b></p>`;
        break;
      case 'code':
        out += `<pre><code>${escapeHtmlChars(tok.text.replace(/\n$/, ''))}</code></pre>`;
        break;
      case 'blockquote':
        out += renderBlocks(tok.tokens ?? []);
        break;
      case 'hr':
        out += '<hr>';
        break;
      case 'space':
        break;
      case 'list':
        out += renderList(tok as Tokens.List);
        break;
      case 'table':
        out += renderTable(tok as Tokens.Table);
        break;
      case 'html':
        // Raw HTML block from the source: strip entirely (never inject).
        break;
      default:
        // Unknown block: render children if any, else drop raw.
        if ('tokens' in tok && Array.isArray((tok as { tokens?: Tokens.Generic[] }).tokens)) {
          out += renderBlocks((tok as { tokens: Tokens.Generic[] }).tokens);
        }
        break;
    }
  }
  return out;
}

/** Renders a list token (recursive for nested lists). */
function renderList(tok: Tokens.List): string {
  const tag = tok.ordered ? 'ol' : 'ul';
  let items = '';
  for (const item of tok.items) {
    let inner = '';
    for (const child of item.tokens) {
      if (child.type === 'text') {
        // 'text' tokens at item level carry the inline content (with 'tokens')
        const t = child as Tokens.Text;
        inner += renderInline(t.tokens ?? [{ type: 'text', raw: t.text }]);
      } else if (child.type === 'list') {
        inner += renderList(child as Tokens.List);
      } else {
        // Block child inside a list item (rare: code, paragraph...)
        inner += renderBlocks([child]);
      }
    }
    items += `<li>${inner}</li>`;
  }
  return `<${tag}>${items}</${tag}>`;
}

/**
 * Renders a table token in the Teams-client shape:
 * bare <tbody>, no <thead>/<th>, every cell wrapped in <p>.
 */
function renderTable(tok: Tokens.Table): string {
  const rows = [tok.header, ...tok.rows];
  const trs = rows
    .map(
      cells =>
        `<tr>${cells
          .map(c => `<td><p>${renderInline(c.tokens)}</p></td>`)
          .join('')}</tr>`
    )
    .join('');
  return `<table><tbody>${trs}</tbody></table>`;
}

/**
 * Converts markdown-formatted text to Teams-compatible HTML.
 * Supports the GFM feature set via marked: bold/italic/strikethrough,
 * inline code, fenced code blocks, nested/ordered/task lists, blockquotes,
 * headings (rendered as bold paragraphs), and tables — all emitted in the
 * shapes the Teams client accepts. Raw HTML in the source is stripped.
 * Plain text without any formatting is wrapped in a paragraph.
 */
export function markdownToTeamsHtml(text: string): string {
  const tokens = md.lexer(text);
  const html = renderBlocks(tokens as Tokens.Generic[]);
  return html || '<p></p>';
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
  // Blockquote, heading, or image
  if (/^\s*>|^\s*#{1,6}\s|!\[[^\]]*\]\(/m.test(text)) return true;
  // Tables (header row + |---| separator)
  if (hasMarkdownTable(text)) return true;
  // Multiple newlines (paragraph breaks)
  if (/\n/.test(text)) return true;

  return false;
}

/** True if the text contains a GFM table (header row followed by a separator row). */
function hasMarkdownTable(text: string): boolean {
  const lines = text.split('\n').map(l => l.trim());
  for (let i = 0; i < lines.length - 1; i++) {
    const isRow = /^\|.*\|$/.test(lines[i]) && lines[i].includes('|', 1);
    const isSep = /^\|[\s:|-]+\|$/.test(lines[i + 1]) && lines[i + 1].includes('-', 1);
    if (isRow && isSep) return true;
  }
  return false;
}
