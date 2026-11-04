/**
 * Profile API - Resolve MRIs to user profiles via fetchShortProfile.
 *
 * Uses the Teams middleTier batch endpoint to resolve MRIs (e.g. "8:orgid:uuid")
 * to display names, job titles, departments, and emails. This is the same API
 * the Teams web client uses to render user cards and reaction tooltips.
 */

import { httpRequest } from '../utils/http.js';
import { MT_API, getSkypeAuthHeaders } from '../utils/api-config.js';
import { type Result, ok } from '../types/result.js';
import { requireSkypeSpacesAuthWithConfig } from '../utils/auth-guards.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/** A resolved user profile from fetchShortProfile. */
export interface ShortProfile {
  mri: string;
  displayName?: string;
  givenName?: string;
  surname?: string;
  email?: string;
  jobTitle?: string;
  department?: string;
  companyName?: string;
  userType?: string;
}

/** Result from batch profile resolution. */
export interface ResolveProfilesResult {
  /** Resolved profiles keyed by MRI. */
  profiles: Map<string, ShortProfile>;
  /** Number of MRIs that could not be resolved. */
  unresolved: number;
}

// ─────────────────────────────────────────────────────────────────────────────
// API Functions
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Resolves MRIs to user profiles via the Teams middleTier fetchShortProfile API.
 *
 * This is a batch endpoint — pass multiple MRIs to resolve them in a single
 * request. Efficient for resolving reactor identities, activity feed senders, etc.
 *
 * @param mris - Array of MRIs to resolve (e.g. ["8:orgid:uuid1", "8:orgid:uuid2"])
 * @returns Map of MRI → profile for each resolved user
 */
export async function resolveProfiles(
  mris: string[]
): Promise<Result<ResolveProfilesResult>> {
  if (mris.length === 0) {
    return ok({ profiles: new Map(), unresolved: 0 });
  }

  const authResult = requireSkypeSpacesAuthWithConfig();
  if (!authResult.ok) {
    return authResult;
  }
  const { skypeToken, spacesToken, regionConfig } = authResult.value;

  const { teamsBaseUrl: baseUrl } = regionConfig;
  const url = MT_API.fetchShortProfile(regionConfig.regionPartition, regionConfig.hasPartition, baseUrl);

  const response = await httpRequest<{ value?: unknown[] }>(
    url,
    {
      method: 'POST',
      headers: getSkypeAuthHeaders(skypeToken, spacesToken, baseUrl),
      body: JSON.stringify(mris),
    }
  );

  if (!response.ok) {
    return response;
  }

  const rawProfiles = response.value.data.value;
  const profiles = new Map<string, ShortProfile>();

  if (Array.isArray(rawProfiles)) {
    for (const raw of rawProfiles) {
      const p = raw as Record<string, unknown>;
      const mri = p.mri as string | undefined;
      if (!mri) continue;

      profiles.set(mri, {
        mri,
        displayName: p.displayName as string | undefined,
        givenName: p.givenName as string | undefined,
        surname: p.surname as string | undefined,
        email: (p.email || p.userPrincipalName) as string | undefined,
        jobTitle: p.jobTitle as string | undefined,
        department: p.department as string | undefined,
        companyName: (p.companyName || p.tenantName) as string | undefined,
        userType: p.userType as string | undefined,
      });
    }
  }

  return ok({
    profiles,
    unresolved: mris.length - profiles.size,
  });
}

/**
 * Resolves MRIs to a simple name map.
 *
 * Convenience wrapper around resolveProfiles() for cases where only
 * display names are needed (e.g. enriching reaction data).
 *
 * @param mris - Array of MRIs to resolve
 * @returns Map of MRI → displayName
 */
export async function resolveNames(
  mris: string[]
): Promise<Map<string, string>> {
  const nameMap = new Map<string, string>();
  if (mris.length === 0) return nameMap;

  try {
    const result = await resolveProfiles(mris);
    if (result.ok) {
      for (const [mri, profile] of result.value.profiles) {
        if (profile.displayName) {
          nameMap.set(mri, profile.displayName);
        }
      }
    }
  } catch {
    // Non-critical: return empty map if resolution fails
  }

  return nameMap;
}
