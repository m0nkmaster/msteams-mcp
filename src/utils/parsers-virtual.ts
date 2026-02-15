/**
 * Virtual conversation message parsing and transcript formatting.
 */

import type { ExtractedLink } from '../types/teams.js';
import { extractLinks, stripHtml } from './parsers-html.js';
import { extractActivityTimestamp, buildMessageLink, type LinkContext } from './parsers-identifiers.js';

// ─────────────────────────────────────────────────────────────────────────────
// Virtual Conversation Parsing
// ─────────────────────────────────────────────────────────────────────────────

/** Common fields from a virtual conversation message (48:saved, 48:threads, etc). */
export interface VirtualConversationItem {
  id: string;
  content: string;
  contentType: string;
  sender: {
    mri: string;
    displayName?: string;
  };
  timestamp: string;
  sourceConversationId: string;
  sourceReferenceId?: string;
  messageLink?: string;
  links?: ExtractedLink[];
}

/**
 * Parses a raw message from a virtual conversation (48:saved, 48:threads, etc).
 * 
 * Virtual conversations contain references to messages in other conversations.
 * The clumpId field contains the source conversation ID, and secondaryReferenceId
 * contains a composite key with the source message/post ID.
 * 
 * @param msg - Raw message object from virtual conversation API
 * @param referencePattern - Regex to extract source ID from secondaryReferenceId
 * @returns Parsed virtual conversation item, or null if message should be skipped
 */
export function parseVirtualConversationMessage(
  msg: Record<string, unknown>,
  referencePattern: RegExp,
  linkContext?: LinkContext
): VirtualConversationItem | null {
  // Skip non-message types
  const messageType = msg.messagetype as string || msg.type as string;
  if (!messageType || messageType.startsWith('Control/')) {
    return null;
  }

  const id = msg.id as string;
  if (!id) return null;

  const content = msg.content as string || '';
  const contentType = messageType || 'Text';

  const fromMri = msg.from as string || '';
  const displayName = msg.imdisplayname as string || msg.displayName as string;

  // Safe timestamp extraction - use extractActivityTimestamp pattern
  const timestamp = extractActivityTimestamp(msg);
  if (!timestamp) return null;

  // clumpId contains the original conversation where the message lives
  const sourceConversationId = msg.clumpId as string || '';
  
  // Extract source reference ID from secondaryReferenceId if available
  let sourceReferenceId: string | undefined;
  const secondaryRef = msg.secondaryReferenceId as string;
  if (secondaryRef) {
    const match = secondaryRef.match(referencePattern);
    if (match) {
      sourceReferenceId = match[1];
    }
  }

  // Build message link to original message
  const messageLink = sourceConversationId && sourceReferenceId
    ? buildMessageLink({
        conversationId: sourceConversationId,
        messageId: sourceReferenceId,
        tenantId: linkContext?.tenantId,
        teamsBaseUrl: linkContext?.teamsBaseUrl,
      })
    : undefined;

  // Extract links before stripping HTML
  const links = extractLinks(content);

  return {
    id,
    content: stripHtml(content),
    contentType,
    sender: {
      mri: fromMri,
      displayName,
    },
    timestamp,
    sourceConversationId,
    sourceReferenceId,
    messageLink,
    links: links.length > 0 ? links : undefined,
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Transcript Formatting
// ─────────────────────────────────────────────────────────────────────────────

/** A single entry from a meeting transcript. */
export interface TranscriptEntry {
  /** Start time (e.g., "00:00:22.287"). */
  startTime: string;
  /** End time (e.g., "00:00:23.167"). */
  endTime: string;
  /** Speaker display name. */
  speaker: string;
  /** Spoken text content. */
  text: string;
}

/**
 * Formats transcript entries into a readable text format.
 * 
 * Merges consecutive entries from the same speaker into a single block
 * to reduce noise and improve readability.
 * 
 * @param entries - Transcript entries
 * @returns Formatted transcript string
 */
export function formatTranscriptText(entries: TranscriptEntry[]): string {
  if (entries.length === 0) return '';
  
  const blocks: string[] = [];
  let currentSpeaker: string | null = null;
  let currentTexts: string[] = [];
  let blockStartTime = '';
  
  for (const entry of entries) {
    if (entry.speaker !== currentSpeaker) {
      // Flush previous block
      if (currentTexts.length > 0) {
        const prefix = currentSpeaker
          ? `[${blockStartTime}] ${currentSpeaker}:`
          : `[${blockStartTime}]`;
        blocks.push(`${prefix}\n${currentTexts.join(' ')}`);
      }
      currentSpeaker = entry.speaker;
      currentTexts = [entry.text];
      blockStartTime = entry.startTime;
    } else {
      currentTexts.push(entry.text);
    }
  }
  
  // Flush last block
  if (currentTexts.length > 0) {
    const prefix = currentSpeaker
      ? `[${blockStartTime}] ${currentSpeaker}:`
      : `[${blockStartTime}]`;
    blocks.push(`${prefix}\n${currentTexts.join(' ')}`);
  }
  
  return blocks.join('\n\n');
}
