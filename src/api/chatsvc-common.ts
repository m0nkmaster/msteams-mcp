/**
 * Shared utilities for chatsvc API sub-modules.
 * 
 * Contains the date formatter and other helpers used across
 * multiple chatsvc sub-modules.
 */

// Reusable date formatter for human-readable timestamps with day of week
// Hoisted to module scope to avoid creating a new formatter per message
const humanReadableDateFormatter = new Intl.DateTimeFormat('en-US', {
  weekday: 'long',
  day: 'numeric',
  month: 'long',
  year: 'numeric',
  hour: '2-digit',
  minute: '2-digit',
  timeZone: 'UTC',
  timeZoneName: 'short',
});

/**
 * Formats an ISO timestamp into a human-readable string with day of week.
 * This helps LLMs correctly identify the day without needing to calculate it.
 * Example: "Friday, January 30, 2026, 10:45 AM UTC"
 */
export function formatHumanReadableDate(isoTimestamp: string): string {
  try {
    const date = new Date(isoTimestamp);
    if (isNaN(date.getTime())) return '';
    return humanReadableDateFormatter.format(date);
  } catch {
    return '';
  }
}
