/**
 * Chat Service API client for messaging operations.
 * 
 * Handles all calls to teams.microsoft.com/api/chatsvc endpoints.
 * Base URL is extracted from session config to support different Teams environments.
 * 
 * This barrel file re-exports from focused sub-modules:
 * - chatsvc-messaging: send, edit, delete, thread messages, 1:1/group chat
 * - chatsvc-activity: activity feed (mentions, reactions, replies)
 * - chatsvc-reactions: add/remove emoji reactions
 * - chatsvc-virtual: saved messages, followed threads, save/unsave
 * - chatsvc-readstatus: consumption horizons, mark as read, unread counts
 * - chatsvc-common: shared utilities (date formatting)
 */

// Re-export everything from sub-modules to maintain backward compatibility
export * from './chatsvc-messaging.js';
export * from './chatsvc-activity.js';
export * from './chatsvc-reactions.js';
export * from './chatsvc-virtual.js';
export * from './chatsvc-readstatus.js';
export * from './chatsvc-common.js';
