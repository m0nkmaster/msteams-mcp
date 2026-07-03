/**
 * HTML parsing and escaping utilities.
 */

import type { ExtractedLink } from '../types/teams.js';

// Re-export ExtractedLink so existing imports from parsers.ts continue to work
export type { ExtractedLink };

/**
 * Extracts links from HTML content before stripping.
 * Returns an array of { url, text, contentType? } objects.
 *
 * Parses both:
 * - <a href="...">text</a> - standard HTML links
 * - <item type="..." uri="..."> - Teams recording/transcript URIs (amsTranscript, onedriveForBusinessVideo, etc.)
 *   Handles type and uri in either order.
 */
export function extractLinks(html: string): ExtractedLink[] {
  const links: ExtractedLink[] = [];

  // <a href="...">text</a> - standard HTML links
  for (const m of html.matchAll(/<a\s+[^>]*href=["'](?<url>[^"']+)["'][^>]*>(?<text>[\s\S]*?)<\/a>/gi)) {
    const { url, text } = m.groups!;
    if (url && !url.startsWith('javascript:')) {
      links.push({ url, text: stripHtml(text) || url });
    }
  }

  // <item type="..." uri="..."> - Teams recording/transcript URIs (attribute order may vary)
  const itemTypeFirst = /<item\s+[^>]*type=["'](?<type>[^"']+)["'][^>]*uri=["'](?<uri>[^"']+)["'][^>]*>/gi;
  const itemUriFirst = /<item\s+[^>]*uri=["'](?<uri>[^"']+)["'][^>]*type=["'](?<type>[^"']+)["'][^>]*>/gi;
  const seen = new Set<string>();
  for (const re of [itemTypeFirst, itemUriFirst]) {
    for (const m of html.matchAll(re)) {
      const { type, uri } = m.groups!;
      if (uri && !seen.has(uri)) {
        seen.add(uri);
        links.push({ url: uri, text: type, contentType: type });
      }
    }
  }

  return links;
}

/**
 * Strips HTML tags from content for display.
 */
export function stripHtml(html: string): string {
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
}

/**
 * Escapes HTML special characters in text.
 */
export function escapeHtmlChars(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Extracts the lowercased URI scheme from a URL, or undefined if the URL is
 * relative (RFC 3986 §3.1).
 *
 * A scheme is the run of letters/digits/`+`/`-`/`.` that starts with a letter
 * and precedes the first `:` — but only when that `:` comes before any `/`, `?`
 * or `#`, so `path/to:thing`, `foo?a=b:c` and `#a:b` are correctly treated as
 * scheme-less relative URLs.
 */
function urlScheme(url: string): string | undefined {
  const colon = url.indexOf(':');
  if (colon <= 0) return undefined;
  const scheme = url.slice(0, colon);
  if (!/^[a-zA-Z][a-zA-Z0-9+.-]*$/.test(scheme)) return undefined;
  // A `/`, `?` or `#` before the colon means the colon is inside a relative URL.
  if (/[/?#]/.test(url.slice(0, colon))) return undefined;
  return scheme.toLowerCase();
}

/**
 * Neutralises a link destination whose URL scheme is not on the allowlist.
 *
 * Only `http`, `https` and `mailto` (case-insensitive) are kept; any other
 * scheme (`javascript:`, `data:`, `vbscript:`, `file:`, …) is rewritten to `#`
 * so a link can never carry an executable scheme into a Teams-rendered message.
 * Scheme-less URLs — relative paths and fragment-only links — carry no
 * executable scheme and pass through unchanged.
 */
export function sanitizeLinkUrl(url: string): string {
  const scheme = urlScheme(url);
  if (scheme !== undefined && !['http', 'https', 'mailto'].includes(scheme)) {
    return '#';
  }
  return url;
}
