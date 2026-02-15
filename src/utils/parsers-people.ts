/**
 * People search result parsing, JWT profile extraction, and token status utilities.
 */

import { extractObjectId } from './parsers-identifiers.js';

/** Person search result from Substrate suggestions API. */
export interface PersonSearchResult {
  id: string;              // Azure AD object ID
  mri: string;             // Teams MRI (8:orgid:guid)
  displayName: string;     // Full display name
  email?: string;          // Primary email address
  givenName?: string;      // First name
  surname?: string;        // Last name
  jobTitle?: string;       // Job title
  department?: string;     // Department
  companyName?: string;    // Company name
}

/** User profile extracted from JWT tokens. */
export interface UserProfile {
  id: string;           // Azure AD object ID (oid)
  mri: string;          // Teams MRI (8:orgid:guid)
  email: string;        // User principal name / email
  displayName: string;  // Full display name
  givenName?: string;   // First name
  surname?: string;     // Last name
  tenantId?: string;    // Azure tenant ID
}

/**
 * Parses a person suggestion from the Substrate API response.
 * 
 * The API can return IDs in various formats:
 * - GUID with tenant: "ab76f827-...@tenant.onmicrosoft.com"
 * - Base64-encoded GUID: "93qkaTtFGWpUHjyRafgdhg=="
 */
export function parsePersonSuggestion(item: Record<string, unknown>): PersonSearchResult | null {
  const rawId = item.Id as string;
  if (!rawId) return null;

  // Extract the ID part (strip tenant suffix if present)
  const idPart = rawId.includes('@') ? rawId.split('@')[0] : rawId;
  
  // Convert to a proper GUID format
  const objectId = extractObjectId(idPart);
  if (!objectId) {
    // If we can't parse the ID, skip this result
    return null;
  }
  
  // Build MRI from the decoded GUID if not provided
  const mri = (item.MRI as string) || `8:orgid:${objectId}`;
  
  const displayName = item.DisplayName as string || '';
  
  // EmailAddresses can be an array
  const emailAddresses = item.EmailAddresses as string[] | undefined;
  const email = emailAddresses?.[0];

  return {
    id: objectId,
    mri: mri.includes('orgid:') && !mri.includes('-') 
      ? `8:orgid:${objectId}`  // Rebuild MRI if it has base64
      : mri,
    displayName,
    email,
    givenName: item.GivenName as string | undefined,
    surname: item.Surname as string | undefined,
    jobTitle: item.JobTitle as string | undefined,
    department: item.Department as string | undefined,
    companyName: item.CompanyName as string | undefined,
  };
}

/**
 * Parses people search results from the Groups/Suggestions structure.
 * 
 * @param groups - Raw Groups array from suggestions API response
 * @returns Array of parsed person results
 */
export function parsePeopleResults(groups: unknown[] | undefined): PersonSearchResult[] {
  const results: PersonSearchResult[] = [];
  
  if (!Array.isArray(groups)) {
    return results;
  }

  for (const group of groups) {
    const g = group as Record<string, unknown>;
    const suggestions = g.Suggestions as unknown[] | undefined;
    
    if (Array.isArray(suggestions)) {
      for (const suggestion of suggestions) {
        const parsed = parsePersonSuggestion(suggestion as Record<string, unknown>);
        if (parsed) results.push(parsed);
      }
    }
  }

  return results;
}

/**
 * Parses user profile from a JWT payload.
 * 
 * @param payload - Decoded JWT payload object
 * @returns User profile or null if required fields are missing
 */
export function parseJwtProfile(payload: Record<string, unknown>): UserProfile | null {
  const oid = payload.oid as string | undefined;
  const name = payload.name as string | undefined;
  
  if (!oid || !name) {
    return null;
  }
  
  const profile: UserProfile = {
    id: oid,
    mri: `8:orgid:${oid}`,
    email: (payload.upn || payload.preferred_username || payload.email || '') as string,
    displayName: name,
    tenantId: payload.tid as string | undefined,
  };
  
  // Try to extract given name and surname
  if (payload.given_name) {
    profile.givenName = payload.given_name as string;
  }
  if (payload.family_name) {
    profile.surname = payload.family_name as string;
  }
  
  // If no given/family name, try to parse from displayName
  if (!profile.givenName && profile.displayName.includes(',')) {
    // Format: "Surname, GivenName"
    const parts = profile.displayName.split(',').map(s => s.trim());
    if (parts.length === 2) {
      profile.surname = parts[0];
      profile.givenName = parts[1];
    }
  } else if (!profile.givenName && profile.displayName.includes(' ')) {
    // Format: "GivenName Surname"
    const parts = profile.displayName.split(' ');
    profile.givenName = parts[0];
    profile.surname = parts.slice(1).join(' ');
  }
  
  return profile;
}

/**
 * Calculates token expiry status from an expiry timestamp.
 * 
 * @param expiryMs - Token expiry time in milliseconds since epoch
 * @param nowMs - Current time in milliseconds (for testing)
 * @returns Token status including whether it's valid and time remaining
 */
export function calculateTokenStatus(
  expiryMs: number,
  nowMs: number = Date.now()
): {
  isValid: boolean;
  expiresAt: string;
  minutesRemaining: number;
} {
  const expiryDate = new Date(expiryMs);
  
  return {
    isValid: expiryMs > nowMs,
    expiresAt: expiryDate.toISOString(),
    minutesRemaining: Math.max(0, Math.round((expiryMs - nowMs) / 1000 / 60)),
  };
}
