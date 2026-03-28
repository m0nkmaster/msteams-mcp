/**
 * Tags API client for Teams channel tag operations.
 *
 * Handles calls to teams.microsoft.com/api/mt/part/{region}/beta/teams/{groupId}/tags endpoints.
 */

import { httpRequest } from '../utils/http.js';
import { TAGS_API, getSkypeAuthHeaders } from '../utils/api-config.js';
import { type Result, ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import { requireSkypeSpacesAuth, getRegionConfig } from '../utils/auth-guards.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/** A Teams channel tag. */
export interface TeamTag {
  /** Tag ID (used in MRI as tag:{id}). */
  id: string;
  /** Display name (e.g., "engineering"). */
  displayName: string;
  /** Number of members assigned to this tag. */
  memberCount: number;
  /** Tag type (e.g., "standard"). */
  tagType: string;
}

/** Result from listing tags. */
export interface ListTagsResult {
  /** Number of tags returned. */
  count: number;
  /** The tags. */
  tags: TeamTag[];
}

// ─────────────────────────────────────────────────────────────────────────────
// API Functions
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Lists all tags for a team.
 *
 * Uses the middleTier API endpoint to fetch tags associated with a team.
 * The teamId (groupId) can be obtained from teams_find_channel results.
 *
 * @param teamId - The team's group ID (GUID)
 * @returns List of tags
 */
export async function listTeamTags(
  teamId: string
): Promise<Result<ListTagsResult>> {
  const authResult = requireSkypeSpacesAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const { skypeToken, spacesToken } = authResult.value;

  const regionConfig = getRegionConfig();
  if (!regionConfig) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'Could not determine region. Please run teams_login to authenticate.',
      { suggestions: ['Call teams_login to authenticate'] }
    ));
  }

  const url = TAGS_API.teamTags(regionConfig.regionPartition, regionConfig.hasPartition, teamId, regionConfig.teamsBaseUrl);

  const response = await httpRequest<Record<string, unknown>>(
    url,
    {
      method: 'GET',
      headers: getSkypeAuthHeaders(skypeToken, spacesToken, regionConfig.teamsBaseUrl),
    }
  );

  if (!response.ok) {
    return response;
  }

  const data = response.value.data;

  const rawTags = (Array.isArray(data) ? data : (data.value as unknown[]) ?? []) as Record<string, unknown>[];

  const tags: TeamTag[] = rawTags.map(tag => ({
    id: (tag.id as string) || '',
    displayName: (tag.displayName as string) || '',
    memberCount: (tag.memberCount as number) || 0,
    tagType: (tag.tagType as string) || 'standard',
  }));

  return ok({
    count: tags.length,
    tags,
  });
}
