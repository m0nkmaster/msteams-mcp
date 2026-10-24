/**
 * Out-of-office (Outlook automatic replies) API client.
 *
 * Reads and writes the mailbox `automaticRepliesSetting` via Microsoft Graph.
 * This is the canonical "Out of Office" that Teams surfaces as the OOF note —
 * Teams reads it from Exchange, so toggling it here flips the status everywhere.
 *
 * Uses the graph.microsoft.com token borrowed from the Teams MSAL cache
 * (delegated scope MailboxSettings.ReadWrite).
 */

import { httpRequest } from '../utils/http.js';
import { type Result, ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import { requireGraphTokenAsync } from '../utils/auth-guards.js';

// ─────────────────────────────────────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────────────────────────────────────

/** Microsoft Graph mailbox settings endpoint (automatic replies live here). */
const MAILBOX_SETTINGS_URL = 'https://graph.microsoft.com/v1.0/me/mailboxSettings';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/** Graph dateTimeTimeZone value. */
export interface GraphDateTime {
  dateTime: string;
  timeZone: string;
}

/** Automatic replies (out-of-office) state as returned by Graph. */
export interface AutomaticRepliesStatus {
  /** "disabled" | "alwaysEnabled" | "scheduled". */
  status: string;
  /** Who receives the external reply: "none" | "contactsOnly" | "all". */
  externalAudience?: string;
  /** Reply sent to senders inside the organisation. */
  internalReplyMessage?: string;
  /** Reply sent to external senders (per externalAudience). */
  externalReplyMessage?: string;
  /** Start of the scheduled window (only when status is "scheduled"). */
  scheduledStartDateTime?: GraphDateTime;
  /** End of the scheduled window (only when status is "scheduled"). */
  scheduledEndDateTime?: GraphDateTime;
}

/** Options for toggling out-of-office. */
export interface SetOutOfOfficeOptions {
  /** true → turn on automatic replies; false → turn off. */
  enabled: boolean;
  /** Reply for internal senders. Reused for external if externalMessage omitted. */
  internalMessage?: string;
  /** Reply for external senders. Defaults to internalMessage. */
  externalMessage?: string;
  /** External audience: "none" | "contactsOnly" | "all". Defaults to "all". */
  externalAudience?: 'none' | 'contactsOnly' | 'all';
  /** ISO 8601 end time. When set (with enabled), uses Graph "scheduled" mode. */
  until?: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Converts an ISO 8601 instant into Graph's UTC dateTimeTimeZone shape.
 */
function toGraphDateTime(iso: string): GraphDateTime {
  return {
    dateTime: new Date(iso).toISOString().replace(/\.\d{3}Z$/, ''),
    timeZone: 'UTC',
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// API Functions
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Gets the current out-of-office (automatic replies) setting.
 */
export async function getAutomaticReplies(): Promise<Result<AutomaticRepliesStatus>> {
  const tokenResult = await requireGraphTokenAsync();
  if (!tokenResult.ok) {
    return tokenResult;
  }

  const response = await httpRequest<{ automaticRepliesSetting?: AutomaticRepliesStatus }>(
    MAILBOX_SETTINGS_URL,
    {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${tokenResult.value}`,
        'Accept': 'application/json',
      },
    }
  );

  if (!response.ok) {
    return response;
  }

  const setting = response.value.data.automaticRepliesSetting;
  if (!setting) {
    return err(createError(
      ErrorCode.API_ERROR,
      'Graph returned no automaticRepliesSetting.',
      { retryable: false }
    ));
  }

  return ok(setting);
}

/**
 * Turns out-of-office on or off.
 *
 * When `enabled` and `until` is set, uses Graph "scheduled" mode (start = now,
 * end = until). When `enabled` without `until`, uses "alwaysEnabled". When not
 * `enabled`, sets "disabled".
 */
export async function setAutomaticReplies(
  options: SetOutOfOfficeOptions
): Promise<Result<AutomaticRepliesStatus>> {
  const tokenResult = await requireGraphTokenAsync();
  if (!tokenResult.ok) {
    return tokenResult;
  }

  const setting: AutomaticRepliesStatus = { status: 'disabled' };

  if (options.enabled) {
    const internal = options.internalMessage ?? '';
    setting.internalReplyMessage = internal;
    setting.externalReplyMessage = options.externalMessage ?? internal;
    setting.externalAudience = options.externalAudience ?? 'all';

    if (options.until) {
      setting.status = 'scheduled';
      setting.scheduledStartDateTime = toGraphDateTime(new Date().toISOString());
      setting.scheduledEndDateTime = toGraphDateTime(options.until);
    } else {
      setting.status = 'alwaysEnabled';
    }
  }

  const response = await httpRequest<{ automaticRepliesSetting?: AutomaticRepliesStatus }>(
    MAILBOX_SETTINGS_URL,
    {
      method: 'PATCH',
      headers: {
        'Authorization': `Bearer ${tokenResult.value}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json',
      },
      body: JSON.stringify({ automaticRepliesSetting: setting }),
    }
  );

  if (!response.ok) {
    return response;
  }

  // Graph echoes the full mailboxSettings; return the updated OOF block
  // (fall back to what we sent if the echo is empty).
  return ok(response.value.data.automaticRepliesSetting ?? setting);
}
