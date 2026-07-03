/**
 * Reaction parsing for chatsvc messages.
 *
 * Parses `properties.emotions` and `annotationsSummary.emotions` from raw
 * chatsvc message payloads into structured reaction data.
 */

import type { RawChatsvcMessage } from '../types/api-responses.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/** A single reaction on a message. */
export interface Reaction {
  /** The emoji key (e.g. "like", "heart", "laugh", or custom "name;storage-id"). */
  key: string;
  /** Who reacted. */
  user: {
    mri: string;
    displayName?: string;
  };
  /** When the reaction was added (epoch ms). */
  time?: number;
}

/** Summary of reactions by emoji key → count. */
export type ReactionSummary = Record<string, number>;

/** Parsed reaction data from a raw message. */
export interface ParsedReactions {
  reactions?: Reaction[];
  reactionSummary?: ReactionSummary;
}

// ─────────────────────────────────────────────────────────────────────────────
// Parser
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Parses reactions from a raw chatsvc message.
 *
 * Handles two data sources:
 * - `properties.emotions`: detailed per-reactor array (present on single-message GET
 *   and list GET with `supportsMessageProperties` view flag)
 * - `annotationsSummary.emotions`: count-per-key summary (often present on both)
 *
 * Returns undefined fields when no reaction data is found, so callers can
 * spread the result without adding empty arrays/objects.
 */
export function parseReactions(msg: RawChatsvcMessage): ParsedReactions {
  const result: ParsedReactions = {};

  // Parse detailed reactions from properties.emotions
  const rawEmotions = msg.properties?.emotions;
  if (rawEmotions) {
    const reactions = parseEmotionsArray(rawEmotions);
    if (reactions.length > 0) {
      result.reactions = reactions;
    }
  }

  // Parse summary from annotationsSummary.emotions
  const summary = (msg as Record<string, unknown>).annotationsSummary;
  if (summary && typeof summary === 'object') {
    const emotions = (summary as Record<string, unknown>).emotions;
    if (emotions && typeof emotions === 'object') {
      const summaryMap: ReactionSummary = {};
      let hasSummary = false;
      for (const [key, count] of Object.entries(emotions as Record<string, unknown>)) {
        if (typeof count === 'number' && count > 0) {
          summaryMap[key] = count;
          hasSummary = true;
        }
      }
      if (hasSummary) {
        result.reactionSummary = summaryMap;
      }
    }
  }

  return result;
}

// ─────────────────────────────────────────────────────────────────────────────
// Internal
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Parses the `properties.emotions` value into structured reactions.
 *
 * The field can be a JSON string or an already-parsed array. Each entry
 * typically has: `{ key, users: [{ mri, time, value }] }` or a flat
 * structure `{ key, mri, time, value }`. We handle both shapes.
 */
function parseEmotionsArray(raw: unknown): Reaction[] {
  let items: unknown[];

  if (typeof raw === 'string') {
    try {
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) return [];
      items = parsed;
    } catch {
      return [];
    }
  } else if (Array.isArray(raw)) {
    items = raw;
  } else {
    return [];
  }

  const reactions: Reaction[] = [];

  for (const item of items) {
    if (!item || typeof item !== 'object') continue;
    const entry = item as Record<string, unknown>;

    const key = entry.key as string | undefined;
    if (!key) continue;

    // Shape 1: { key, users: [{ mri, time, value }] }
    const users = entry.users;
    if (Array.isArray(users)) {
      for (const u of users) {
        if (!u || typeof u !== 'object') continue;
        const user = u as Record<string, unknown>;
        reactions.push({
          key,
          user: {
            mri: String(user.mri || ''),
            displayName: user.displayName as string | undefined,
          },
          time: typeof user.time === 'number' ? user.time
            : typeof user.value === 'number' ? user.value
            : undefined,
        });
      }
      continue;
    }

    // Shape 2: flat { key, mri, time/value }
    if (entry.mri) {
      reactions.push({
        key,
        user: {
          mri: String(entry.mri),
          displayName: entry.displayName as string | undefined,
        },
        time: typeof entry.time === 'number' ? entry.time
          : typeof entry.value === 'number' ? entry.value
          : undefined,
      });
    }
  }

  return reactions;
}
