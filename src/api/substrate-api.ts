/**
 * Substrate API client for search and people operations.
 * 
 * Handles all calls to substrate.office.com endpoints.
 */

import { httpRequest } from '../utils/http.js';
import { SUBSTRATE_API, getBearerHeaders } from '../utils/api-config.js';
import { ErrorCode } from '../types/errors.js';
import { type Result, ok } from '../types/result.js';
import { clearTokenCache } from '../auth/token-extractor.js';
import { requireSubstrateTokenAsync } from '../utils/auth-guards.js';
import {
  parseSearchResults,
  parsePeopleResults,
  parseChannelResults,
  filterChannelsByName,
  type PersonSearchResult,
  type ChannelSearchResult,
} from '../utils/parsers.js';
import { getMyTeamsAndChannels } from './csa-api.js';
import type { TeamsSearchResult, SearchPaginationResult } from '../types/teams.js';

/** Search result with pagination. */
export interface SearchResult {
  results: TeamsSearchResult[];
  pagination: SearchPaginationResult;
}

/** People search result. */
export interface PeopleSearchResult {
  results: PersonSearchResult[];
  returned: number;
}

/**
 * Searches Teams messages using the Substrate v2 query API.
 */
export async function searchMessages(
  query: string,
  options: { from?: number; size?: number; maxResults?: number } = {}
): Promise<Result<SearchResult>> {
  const tokenResult = await requireSubstrateTokenAsync();
  if (!tokenResult.ok) {
    return tokenResult;
  }
  const token = tokenResult.value;

  const from = options.from ?? 0;
  const size = options.size ?? 25;

  // Generate unique IDs for this request
  const cvid = crypto.randomUUID();
  const logicalId = crypto.randomUUID();

  const body = {
    entityRequests: [{
      entityType: 'Message',
      contentSources: ['Teams'],
      propertySet: 'Optimized',
      fields: [
        'Extension_SkypeSpaces_ConversationPost_Extension_FromSkypeInternalId_String',
        'Extension_SkypeSpaces_ConversationPost_Extension_ThreadType_String',
        'Extension_SkypeSpaces_ConversationPost_Extension_SkypeGroupId_String',
      ],
      query: {
        queryString: `${query} AND NOT (isClientSoftDeleted:TRUE)`,
        displayQueryString: query,
      },
      from,
      size,
      topResultsCount: 5,
    }],
    QueryAlterationOptions: {
      EnableAlteration: true,
      EnableSuggestion: true,
      SupportedRecourseDisplayTypes: ['Suggestion'],
    },
    cvid,
    logicalId,
    scenario: {
      Dimensions: [
        { DimensionName: 'QueryType', DimensionValue: 'Messages' },
        { DimensionName: 'FormFactor', DimensionValue: 'general.web.reactSearch' },
      ],
      Name: 'powerbar',
    },
    timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
  };

  const response = await httpRequest<Record<string, unknown>>(
    SUBSTRATE_API.search,
    {
      method: 'POST',
      headers: getBearerHeaders(token),
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    // Clear cache on auth errors
    if (response.error.code === ErrorCode.AUTH_EXPIRED) {
      clearTokenCache();
    }
    return response;
  }

  const data = response.value.data;
  const { results, total } = parseSearchResults(
    data.EntitySets as unknown[] | undefined
  );

  const maxResults = options.maxResults ?? size;
  const limitedResults = results.slice(0, maxResults);

  return ok({
    results: limitedResults,
    pagination: {
      from,
      size,
      returned: limitedResults.length,
      total,
      hasMore: total !== undefined
        ? from + results.length < total
        : results.length >= size,
    },
  });
}

/**
 * Searches for people by name or email.
 */
export async function searchPeople(
  query: string,
  limit: number = 10
): Promise<Result<PeopleSearchResult>> {
  const tokenResult = await requireSubstrateTokenAsync();
  if (!tokenResult.ok) {
    return tokenResult;
  }
  const token = tokenResult.value;

  const cvid = crypto.randomUUID();
  const logicalId = crypto.randomUUID();

  const body = {
    EntityRequests: [{
      Query: {
        QueryString: query,
        DisplayQueryString: query,
      },
      EntityType: 'People',
      Size: limit,
      Fields: [
        'Id',
        'MRI',
        'DisplayName',
        'EmailAddresses',
        'GivenName',
        'Surname',
        'JobTitle',
        'Department',
        'CompanyName',
      ],
    }],
    cvid,
    logicalId,
  };

  const response = await httpRequest<Record<string, unknown>>(
    SUBSTRATE_API.peopleSearch,
    {
      method: 'POST',
      headers: getBearerHeaders(token),
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    if (response.error.code === ErrorCode.AUTH_EXPIRED) {
      clearTokenCache();
    }
    return response;
  }

  const results = parsePeopleResults(response.value.data.Groups as unknown[] | undefined);

  return ok({
    results,
    returned: results.length,
  });
}

/**
 * Gets the user's frequently contacted people.
 */
export async function getFrequentContacts(
  limit: number = 50
): Promise<Result<PeopleSearchResult>> {
  const tokenResult = await requireSubstrateTokenAsync();
  if (!tokenResult.ok) {
    return tokenResult;
  }
  const token = tokenResult.value;

  const cvid = crypto.randomUUID();
  const logicalId = crypto.randomUUID();

  const body = {
    EntityRequests: [{
      Query: {
        QueryString: '',
        DisplayQueryString: '',
      },
      EntityType: 'People',
      Size: limit,
      Fields: [
        'Id',
        'MRI',
        'DisplayName',
        'EmailAddresses',
        'GivenName',
        'Surname',
        'JobTitle',
        'Department',
        'CompanyName',
      ],
    }],
    cvid,
    logicalId,
  };

  const response = await httpRequest<Record<string, unknown>>(
    SUBSTRATE_API.frequentContacts,
    {
      method: 'POST',
      headers: getBearerHeaders(token),
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    if (response.error.code === ErrorCode.AUTH_EXPIRED) {
      clearTokenCache();
    }
    return response;
  }

  const contacts = parsePeopleResults(response.value.data.Groups as unknown[] | undefined);

  return ok({
    results: contacts,
    returned: contacts.length,
  });
}

/** Channel search result. */
export interface ChannelSearchResultSet {
  results: ChannelSearchResult[];
  returned: number;
}

/**
 * Searches for Teams channels by name using both:
 * 1. User's own teams/channels (Teams List API) - reliable, shows membership
 * 2. Organisation-wide discovery (Substrate suggestions) - broader but less reliable
 * 
 * Results are merged and deduplicated, with membership status indicated.
 * 
 * @param query - Channel name to search for
 * @param limit - Maximum number of results (default: 10, max: 50)
 */
export async function searchChannels(
  query: string,
  limit: number = 10
): Promise<Result<ChannelSearchResultSet>> {
  const tokenResult = await requireSubstrateTokenAsync();
  if (!tokenResult.ok) {
    return tokenResult;
  }
  const token = tokenResult.value;

  // Run both searches in parallel
  const [orgSearchResult, myTeamsResult] = await Promise.all([
    searchChannelsOrgWide(query, limit, token),
    getMyTeamsAndChannels(),
  ]);

  // Build a map of channel IDs the user is a member of
  const memberChannelIds = new Set<string>();
  const myChannelsMatching: ChannelSearchResult[] = [];

  if (myTeamsResult.ok) {
    // Filter the user's channels by the query and collect matching ones
    const matching = filterChannelsByName(myTeamsResult.value.teams, query);
    for (const channel of matching) {
      memberChannelIds.add(channel.channelId);
      myChannelsMatching.push(channel);
    }
    
    // Also add all channel IDs to the set for membership lookup
    for (const team of myTeamsResult.value.teams) {
      for (const channel of team.channels) {
        memberChannelIds.add(channel.channelId);
      }
    }
  }

  // Process org-wide results, marking membership status
  const orgChannels: ChannelSearchResult[] = [];
  if (orgSearchResult.ok) {
    for (const channel of orgSearchResult.value) {
      // Mark whether user is a member
      channel.isMember = memberChannelIds.has(channel.channelId);
      orgChannels.push(channel);
    }
  }

  // Merge results: start with user's matching channels (definitely accessible),
  // then add org-wide results that aren't duplicates
  const seenIds = new Set<string>();
  const merged: ChannelSearchResult[] = [];

  // First add channels from user's teams (reliable, known accessible)
  for (const channel of myChannelsMatching) {
    if (!seenIds.has(channel.channelId)) {
      seenIds.add(channel.channelId);
      merged.push(channel);
    }
  }

  // Then add org-wide results that aren't duplicates
  for (const channel of orgChannels) {
    if (!seenIds.has(channel.channelId)) {
      seenIds.add(channel.channelId);
      merged.push(channel);
    }
  }

  // Apply limit
  const limited = merged.slice(0, limit);

  return ok({
    results: limited,
    returned: limited.length,
  });
}

/**
 * Internal: Searches for channels org-wide using the Substrate suggestions API.
 * 
 * This is a typeahead/autocomplete API, so matching behaviour may be inconsistent
 * for multi-word queries. Used as a supplement to the user's own teams list.
 */
async function searchChannelsOrgWide(
  query: string,
  limit: number,
  token: string
): Promise<Result<ChannelSearchResult[]>> {
  const cvid = crypto.randomUUID();
  const logicalId = crypto.randomUUID();

  const body = {
    EntityRequests: [{
      Query: {
        QueryString: query,
        DisplayQueryString: query,
      },
      EntityType: 'TeamsChannel',
      Size: Math.min(limit, 50),
    }],
    cvid,
    logicalId,
  };

  const response = await httpRequest<Record<string, unknown>>(
    SUBSTRATE_API.channelSearch,
    {
      method: 'POST',
      headers: getBearerHeaders(token),
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    if (response.error.code === ErrorCode.AUTH_EXPIRED) {
      clearTokenCache();
    }
    return response;
  }

  const data = response.value.data as Record<string, unknown> | undefined;
  const results = parseChannelResults(data?.Groups as unknown[] | undefined);

  // Mark all org-wide results as isMember: false initially
  // (caller will update based on actual membership)
  for (const result of results) {
    result.isMember = false;
  }

  return ok(results);
}
