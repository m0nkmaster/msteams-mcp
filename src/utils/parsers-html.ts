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
