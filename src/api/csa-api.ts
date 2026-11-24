/**
 * CSA (Chat Service Aggregator) API client for favorites and teams operations.
 * 
 * Handles all calls to teams.microsoft.com/api/csa endpoints.
 * Base URL is extracted from session config to support different Teams environments.
 */

import { httpRequest } from '../utils/http.js';
import { CSA_API, getCsaHeaders } from '../utils/api-config.js';
import { type Result, ok } from '../types/result.js';
import { requireCsaAuth, getApiConfig } from '../utils/auth-guards.js';
import { getFolderHierarchy, enrichSectionItems, updateSection } from './sections-api.js';
import {
  parseTeamsList,
  type TeamWithChannels,
} from '../utils/parsers.js';

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
  const state = await getFolderHierarchy();
  if (!state.ok) return state;
  const folder = state.value.conversationFolders.find(f => f.folderType === 'Favorites');
  const items = (folder?.conversationFolderItems ?? []).flatMap(item =>
    item.conversationId ? [{ ...item, conversationId: item.conversationId }] : []);
  return ok({
    favorites: await enrichSectionItems(items),
    folderHierarchyVersion: state.value.folderHierarchyVersion,
    folderId: folder?.id,
  });
}

/**
 * Adds a conversation to the user's favourites.
 */
export async function addFavorite(
  conversationId: string
): Promise<Result<void>> {
  const result = await updateSection({ action: 'add', conversationId });
  return result.ok ? ok(undefined) : result;
}

/** Removes a conversation from the built-in Favorites section. */
export async function removeFavorite(conversationId: string): Promise<Result<void>> {
  const result = await updateSection({ action: 'remove', conversationId });
  return result.ok ? ok(undefined) : result;
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

  const { csaRegion, baseUrl } = getApiConfig();
  const url = CSA_API.teamsList(csaRegion, baseUrl);

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

  const { csaRegion, baseUrl } = getApiConfig();
  const url = CSA_API.customEmojis(csaRegion, baseUrl);

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
