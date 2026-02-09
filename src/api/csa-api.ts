/**
 * CSA (Chat Service Aggregator) API client for favorites and teams operations.
 * 
 * Handles all calls to teams.microsoft.com/api/csa endpoints.
 * Base URL is extracted from session config to support different Teams environments.
 */

import { httpRequest } from '../utils/http.js';
import { CSA_API, getCsaHeaders } from '../utils/api-config.js';
import { ErrorCode, createError } from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';
import { requireCsaAuth, getRegion, getTeamsBaseUrl } from '../utils/auth-guards.js';
import {
  getConversationProperties,
  extractParticipantNames,
} from './chatsvc-api.js';
import {
  parseTeamsList,
  type TeamWithChannels,
} from '../utils/parsers.js';

/** Gets region and base URL together for API calls. */
function getApiConfig() {
  return {
    region: getRegion(),
    baseUrl: getTeamsBaseUrl(),
  };
}

/** A favourite/pinned conversation item. */
export interface FavoriteItem {
  conversationId: string;
  displayName?: string;
  conversationType?: string;
  createdTime?: number;
  lastUpdatedTime?: number;
}

/** Response from getting favorites. */
export interface FavoritesResult {
  favorites: FavoriteItem[];
  folderHierarchyVersion?: number;
  folderId?: string;
}

/**
 * Gets the user's favourite/pinned conversations.
 */
export async function getFavorites(): Promise<Result<FavoritesResult>> {
  const authResult = requireCsaAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, csaToken } = authResult.value;

  const { region, baseUrl } = getApiConfig();
  const url = CSA_API.conversationFolders(region, baseUrl);

  const response = await httpRequest<Record<string, unknown>>(
    url,
    {
      method: 'GET',
      headers: getCsaHeaders(auth.skypeToken, csaToken, baseUrl),
    }
  );

  if (!response.ok) {
    return response;
  }

  const data = response.value.data;

  // Find the Favorites folder
  const folders = data.conversationFolders as unknown[] | undefined;
  const favoritesFolder = folders?.find((f: unknown) => {
    const folder = f as Record<string, unknown>;
    return folder.folderType === 'Favorites';
  }) as Record<string, unknown> | undefined;

  if (!favoritesFolder) {
    return ok({
      favorites: [],
      folderHierarchyVersion: data.folderHierarchyVersion as number,
    });
  }

  const items = favoritesFolder.conversationFolderItems as unknown[] | undefined;
  const favorites: FavoriteItem[] = (items || []).map((item: unknown) => {
    const i = item as Record<string, unknown>;
    return {
      conversationId: i.conversationId as string,
      createdTime: i.createdTime as number | undefined,
      lastUpdatedTime: i.lastUpdatedTime as number | undefined,
    };
  });

  // Enrich favorites with display names in parallel
  const enrichmentPromises = favorites.map(async (fav) => {
    const props = await getConversationProperties(fav.conversationId);
    if (props.ok) {
      fav.displayName = props.value.displayName;
      fav.conversationType = props.value.conversationType;
    }

    // Fallback: extract from recent messages if no display name
    if (!fav.displayName) {
      const names = await extractParticipantNames(fav.conversationId);
      if (names.ok && names.value) {
        fav.displayName = names.value;
      }
    }
  });

  await Promise.allSettled(enrichmentPromises);

  return ok({
    favorites,
    folderHierarchyVersion: data.folderHierarchyVersion as number,
    folderId: favoritesFolder.id as string,
  });
}

/**
 * Adds a conversation to the user's favourites.
 */
export async function addFavorite(
  conversationId: string
): Promise<Result<void>> {
  return modifyFavorite(conversationId, 'AddItem');
}

/**
 * Removes a conversation from the user's favourites.
 */
export async function removeFavorite(
  conversationId: string
): Promise<Result<void>> {
  return modifyFavorite(conversationId, 'RemoveItem');
}

/**
 * Internal helper to modify the favourites folder.
 */
async function modifyFavorite(
  conversationId: string,
  action: 'AddItem' | 'RemoveItem'
): Promise<Result<void>> {
  const authResult = requireCsaAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, csaToken } = authResult.value;

  const { region, baseUrl } = getApiConfig();

  // Get current folder state
  const currentState = await getFavorites();
  if (!currentState.ok) {
    return err(currentState.error);
  }

  if (!currentState.value.folderId) {
    return err(createError(
      ErrorCode.NOT_FOUND,
      'Could not find Favorites folder'
    ));
  }

  const url = CSA_API.conversationFolders(region, baseUrl);

  const response = await httpRequest<unknown>(
    url,
    {
      method: 'POST',
      headers: getCsaHeaders(auth.skypeToken, csaToken, baseUrl),
      body: JSON.stringify({
        folderHierarchyVersion: currentState.value.folderHierarchyVersion,
        actions: [
          {
            action,
            folderId: currentState.value.folderId,
            itemId: conversationId,
          },
        ],
      }),
    }
  );

  if (!response.ok) {
    return response;
  }

  return ok(undefined);
}

/** Response from getting the user's teams and channels. */
export interface TeamsListResult {
  teams: TeamWithChannels[];
}

/**
 * Gets all teams and channels the user is a member of.
 * 
 * This returns the complete list of teams with their channels - not a search,
 * but a full enumeration of the user's memberships.
 */
export async function getMyTeamsAndChannels(): Promise<Result<TeamsListResult>> {
  const authResult = requireCsaAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, csaToken } = authResult.value;

  const { region, baseUrl } = getApiConfig();
  const url = CSA_API.teamsList(region, baseUrl);

  const response = await httpRequest<Record<string, unknown>>(
    url,
    {
      method: 'GET',
      headers: getCsaHeaders(auth.skypeToken, csaToken, baseUrl),
    }
  );

  if (!response.ok) {
    return response;
  }

  const teams = parseTeamsList(response.value.data);

  return ok({ teams });
}

// ─────────────────────────────────────────────────────────────────────────────
// Custom Emoji Operations
// ─────────────────────────────────────────────────────────────────────────────

/** A custom emoji from the organisation. */
export interface CustomEmoji {
  /** The emoji ID (use as reaction key). */
  id: string;
  /** Short name/shortcut for the emoji. */
  shortcut: string;
  /** Description of the emoji. */
  description: string;
  /** When the emoji was created. */
  createdOn?: number;
}

/** Response from getting custom emojis. */
export interface CustomEmojisResult {
  emojis: CustomEmoji[];
}

/**
 * Gets the organisation's custom emojis.
 */
export async function getCustomEmojis(): Promise<Result<CustomEmojisResult>> {
  const authResult = requireCsaAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, csaToken } = authResult.value;

  const { region, baseUrl } = getApiConfig();
  const url = CSA_API.customEmojis(region, baseUrl);

  const response = await httpRequest<Record<string, unknown>>(
    url,
    {
      method: 'GET',
      headers: getCsaHeaders(auth.skypeToken, csaToken, baseUrl),
    }
  );

  if (!response.ok) {
    return response;
  }

  const data = response.value.data;
  const categories = data.categories as Array<Record<string, unknown>> | undefined;
  
  const emojis: CustomEmoji[] = [];
  
  if (categories) {
    for (const category of categories) {
      const emoticons = category.emoticons as Array<Record<string, unknown>> | undefined;
      if (emoticons) {
        for (const emoticon of emoticons) {
          if (emoticon.isDeleted) continue;
          
          const id = emoticon.id as string;
          const shortcuts = emoticon.shortcuts as string[] | undefined;
          const description = emoticon.description as string | undefined;
          const createdOn = emoticon.createdOn as number | undefined;
          
          emojis.push({
            id,
            shortcut: shortcuts?.[0] || id.split(';')[0],
            description: description || shortcuts?.[0] || id.split(';')[0],
            createdOn,
          });
        }
      }
    }
  }

  return ok({ emojis });
}
