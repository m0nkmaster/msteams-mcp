/**
 * Presence API client.
 *
 * Fetches availability/activity, out of office state, and auto reply notes for
 * one or more users from the Teams User Presence Service (UPS).
 */

import { httpRequest } from '../utils/http.js';
import { PRESENCE_API, getBearerHeaders } from '../utils/api-config.js';
import { type Result, ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import { requireSkypeSpacesAuth, getRegionConfig } from '../utils/auth-guards.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/** Presence and out of office information for a single user. */
export interface UserPresence {
  /** The user's MRI (e.g. "8:orgid:<guid>"). */
  id: string;
  /** Availability status (e.g. Available, Busy, Away, Offline, DoNotDisturb). */
  availability: string;
  /** Fine grained activity (e.g. Available, InACall, InAMeeting, Presenting, Offline). */
  activity: string;
  /** Device the user is active on (e.g. Web, Mobile, Desktop), if known. */
  deviceType?: string;
  /** Last time the user was active (ISO 8601), if known. */
  lastActiveTime?: string;
  /** Whether the user is currently out of office per their calendar. */
  isOutOfOffice: boolean;
  /** The user's out of office / auto reply note, if set. */
  outOfOfficeMessage?: string;
  /** When the out of office note expires (ISO 8601), if set. */
  outOfOfficeExpiry?: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Normalises a user identifier to an MRI.
 *
 * Accepts an MRI ("8:orgid:<guid>"), an object id with tenant ("<guid>@<tenant>"),
 * or a raw object id ("<guid>"), and returns the MRI form used by the presence API.
 */
function toMri(userId: string): string {
  if (userId.startsWith('8:') || userId.startsWith('28:')) {
    return userId;
  }
  const guid = userId.includes('@') ? userId.split('@')[0] : userId;
  return `8:orgid:${guid}`;
}

/**
 * Parses a raw UPS presence entry into our UserPresence type.
 */
function parsePresence(raw: Record<string, unknown>): UserPresence {
  const presence = (raw.presence ?? {}) as Record<string, unknown>;
  const calendarData = (presence.calendarData ?? {}) as Record<string, unknown>;
  const oofNote = calendarData.outOfOfficeNote as
    | { message?: string; expiry?: string }
    | undefined;

  const message = oofNote?.message?.trim();

  return {
    id: raw.mri as string,
    availability: (presence.availability as string) ?? 'Unknown',
    activity: (presence.activity as string) ?? 'Unknown',
    deviceType: presence.deviceType as string | undefined,
    lastActiveTime: presence.lastActiveTime as string | undefined,
    isOutOfOffice: calendarData.isOutOfOffice === true,
    outOfOfficeMessage: message ? message : undefined,
    outOfOfficeExpiry: oofNote?.expiry,
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// API Functions
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Gets presence (and out of office / auto reply notes) for one or more users.
 *
 * @param userIds - User MRIs, object ids with tenant, or raw object ids
 * @returns Presence information for each requested user
 */
export async function getPresence(
  userIds: string[]
): Promise<Result<UserPresence[]>> {
  const authResult = requireSkypeSpacesAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const { spacesToken } = authResult.value;

  const regionConfig = getRegionConfig();
  if (!regionConfig) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'Could not determine region. Please run teams_login to authenticate.',
      { suggestions: ['Call teams_login to authenticate'] }
    ));
  }

  const body = userIds.map((id) => ({ mri: toMri(id), source: 'ups' }));

  // The presence (ups) service uses the geo region (e.g. "emea"), which is the
  // prefix of the partitioned region ("emea-02"), not the chatsvc country code.
  const geoRegion = regionConfig.regionPartition.split('-')[0];
  const url = PRESENCE_API.getPresence(geoRegion, regionConfig.teamsBaseUrl);

  const response = await httpRequest<Array<Record<string, unknown>>>(
    url,
    {
      method: 'POST',
      headers: getBearerHeaders(spacesToken, regionConfig.teamsBaseUrl),
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    return response;
  }

  const data = response.value.data;
  const entries = Array.isArray(data) ? data : [];

  return ok(entries.map(parsePresence));
}
