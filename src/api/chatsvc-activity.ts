/**
 * Chat Service API - Activity feed operations.
 * 
 * Fetches the user's activity feed (mentions, reactions, replies, notifications).
 */

import { httpRequest } from '../utils/http.js';
import { CHATSVC_API, getSkypeAuthHeaders } from '../utils/api-config.js';
import { type Result, ok } from '../types/result.js';
import { requireMessageAuthWithConfig, getTeamsBaseUrl, getTenantId } from '../utils/auth-guards.js';
import { stripHtml, extractLinks, buildMessageLink, extractActivityTimestamp, type ExtractedLink } from '../utils/parsers.js';
import { DEFAULT_ACTIVITY_LIMIT, VIRTUAL_CONVERSATION_PREFIX } from '../constants.js';
import type { RawChatsvcMessage } from '../types/api-responses.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/** Type of activity item. */
export type ActivityType = 'mention' | 'reaction' | 'reply' | 'message' | 'unknown';

/** An activity item from the notifications feed. */
export interface ActivityItem {
  /** Activity/message ID. */
  id: string;
  /** Type of activity (mention, reaction, reply, etc.). */
  type: ActivityType;
  /** Activity content (HTML stripped). */
  content: string;
  /** Raw content type from API. */
  contentType: string;
  /** Sender information. */
  sender: {
    mri: string;
    displayName?: string;
  };
  /** When the activity occurred. */
  timestamp: string;
  /** The source conversation ID (where the activity happened). */
  conversationId?: string;
  /** The conversation/thread topic name. */
  topic?: string;
  /** Direct link to the activity in Teams. */
  activityLink?: string;
  /** Links extracted from content. */
  links?: ExtractedLink[];
}

/** Result of fetching the activity feed. */
export interface GetActivityResult {
  /** Activity items (newest first). */
  activities: ActivityItem[];
  /** Sync state for incremental polling. */
  syncState?: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// Internal Helpers
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Determines the activity type from message content and properties.
 */
function detectActivityType(msg: Record<string, unknown>): ActivityType {
  const content = msg.content as string || '';
  const messageType = msg.messagetype as string || '';
  
  // Check for @mention
  if (content.includes('itemtype="http://schema.skype.com/Mention"') ||
      content.includes('itemscope itemtype="http://schema.skype.com/Mention"')) {
    return 'mention';
  }
  
  // Check for reaction-related message types
  if (messageType.toLowerCase().includes('reaction')) {
    return 'reaction';
  }
  
  // Check for thread/reply indicators
  if (msg.threadtopic || msg.parentMessageId) {
    return 'reply';
  }
  
  // Standard message
  if (messageType.includes('RichText') || messageType.includes('Text')) {
    return 'message';
  }
  
  return 'unknown';
}

// ─────────────────────────────────────────────────────────────────────────────
// API Functions
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Gets the activity feed (notifications) for the current user.
 * Includes mentions, reactions, replies, and other notifications.
 */
export async function getActivityFeed(
  options: { limit?: number } = {}
): Promise<Result<GetActivityResult>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, region, baseUrl } = authResult.value;
  const limit = options.limit ?? DEFAULT_ACTIVITY_LIMIT;

  let url = CHATSVC_API.activityFeed(region, baseUrl);
  url += `?view=msnp24Equivalent&pageSize=${limit}`;

  const response = await httpRequest<{ messages?: unknown[]; syncState?: string }>(
    url,
    {
      method: 'GET',
      headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken, baseUrl),
    }
  );

  if (!response.ok) {
    return response;
  }

  const rawMessages = response.value.data.messages;
  const syncState = response.value.data.syncState;
  
  if (!Array.isArray(rawMessages)) {
    return ok({
      activities: [],
      syncState,
    });
  }

  const activities: ActivityItem[] = [];

  for (const raw of rawMessages) {
    const msg = raw as RawChatsvcMessage;

    // Skip control/system messages that aren't relevant
    const messageType = msg.messagetype;
    if (!messageType || 
        messageType.startsWith('Control/') || 
        messageType === 'ThreadActivity/AddMember' ||
        messageType === 'ThreadActivity/DeleteMember') {
      continue;
    }

    const id = msg.id || msg.originalarrivaltime;
    if (!id) continue;

    const content = msg.content || '';
    const contentType = msg.messagetype || 'Text';

    const fromMri = msg.from || '';
    const displayName = msg.imdisplayname || msg.displayName;

    // Safely extract timestamp - returns null if no valid timestamp found
    const timestamp = extractActivityTimestamp(msg);
    if (!timestamp) continue;

    // Get source conversation - prefer clumpId (actual source) over conversationid
    // Some activity items have conversationid as "48:notifications" (the virtual conversation)
    // which doesn't work for deep links. clumpId contains the real source conversation.
    const rawConversationId = msg.conversationid || msg.conversationId;
    const clumpId = msg.clumpId;
    
    // Use clumpId if conversationid is a virtual conversation (48:xxx format)
    const isVirtualConversation = rawConversationId?.startsWith(VIRTUAL_CONVERSATION_PREFIX);
    const conversationId = (isVirtualConversation && clumpId) ? clumpId : rawConversationId;
    
    const topic = msg.threadtopic || msg.topic;

    // Build activity link if we have a valid source conversation context
    // Skip virtual conversations (48:xxx) as they don't produce working deep links
    let activityLink: string | undefined;
    if (conversationId && !conversationId.startsWith(VIRTUAL_CONVERSATION_PREFIX) && /^\d+$/.test(id)) {
      activityLink = buildMessageLink({
        conversationId,
        messageId: id,
        tenantId: getTenantId() ?? undefined,
        teamsBaseUrl: getTeamsBaseUrl(),
      });
    }

    const activityType = detectActivityType(msg);

    // Extract links before stripping HTML
    const links = extractLinks(content);

    activities.push({
      id,
      type: activityType,
      content: stripHtml(content),
      contentType,
      sender: {
        mri: fromMri,
        displayName,
      },
      timestamp,
      conversationId,
      topic,
      activityLink,
      links: links.length > 0 ? links : undefined,
    });
  }

  // Sort by timestamp (newest first for activity feed)
  activities.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());

  return ok({
    activities,
    syncState,
  });
}
