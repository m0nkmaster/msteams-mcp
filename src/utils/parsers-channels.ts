/**
 * Channel and team list parsing from Substrate and CSA API responses.
 */

/** Channel search result from Substrate suggestions API or Teams List API. */
export interface ChannelSearchResult {
  channelId: string;         // Conversation ID (19:xxx@thread.tacv2)
  channelName: string;       // Channel display name
  teamName: string;          // Parent team name
  teamId: string;            // Team group ID
  channelType: string;       // "Standard", "Private", etc.
  description?: string;      // Channel description if available
  isMember?: boolean;        // Whether user is a member of this channel's team
}

/**
 * Parses a single channel suggestion from the API response.
 * 
 * @param suggestion - Raw suggestion object from API
 * @returns Parsed channel result or null if required fields are missing
 */
export function parseChannelSuggestion(
  suggestion: Record<string, unknown>
): ChannelSearchResult | null {
  const name = suggestion.Name as string | undefined;
  const threadId = suggestion.ThreadId as string | undefined;
  const teamName = suggestion.TeamName as string | undefined;
  const groupId = suggestion.GroupId as string | undefined;
  
  // All required fields must be present
  if (!name || !threadId || !teamName || !groupId) {
    return null;
  }

  return {
    channelId: threadId,
    channelName: name,
    teamName,
    teamId: groupId,
    channelType: (suggestion.ChannelType as string) || 'Standard',
    description: suggestion.Description as string | undefined,
  };
}

/**
 * Parses channel search results from the Groups/Suggestions structure.
 * 
 * @param groups - Raw Groups array from suggestions API response
 * @returns Array of parsed channel results
 */
export function parseChannelResults(groups: unknown[] | undefined): ChannelSearchResult[] {
  const results: ChannelSearchResult[] = [];
  
  if (!Array.isArray(groups)) {
    return results;
  }

  for (const group of groups) {
    const g = group as Record<string, unknown>;
    const suggestions = g.Suggestions as unknown[] | undefined;
    
    if (Array.isArray(suggestions)) {
      for (const suggestion of suggestions) {
        const s = suggestion as Record<string, unknown>;
        // Only parse ChannelSuggestion entities
        if (s.EntityType === 'ChannelSuggestion') {
          const parsed = parseChannelSuggestion(s);
          if (parsed) results.push(parsed);
        }
      }
    }
  }

  return results;
}

/** Team with channels from the Teams List API response. */
export interface TeamWithChannels {
  teamId: string;           // Team group ID (GUID)
  teamName: string;         // Team display name
  threadId: string;         // Team root conversation ID
  description?: string;     // Team description
  channels: ChannelSearchResult[];
}

/**
 * Parses the Teams List API response to extract all teams and channels.
 * 
 * @param data - Raw response data from /api/csa/{region}/api/v3/teams/users/me
 * @returns Array of teams with their channels
 */
export function parseTeamsList(data: Record<string, unknown> | undefined): TeamWithChannels[] {
  const results: TeamWithChannels[] = [];
  
  if (!data) return results;
  
  const teams = data.teams as unknown[] | undefined;
  if (!Array.isArray(teams)) return results;
  
  for (const team of teams) {
    const t = team as Record<string, unknown>;
    // Team's id IS the thread ID (format: 19:xxx@thread.tacv2)
    const threadId = t.id as string | undefined;
    const displayName = t.displayName as string | undefined;
    
    if (!threadId || !displayName) continue;
    
    const channels: ChannelSearchResult[] = [];
    const channelList = t.channels as unknown[] | undefined;
    
    if (Array.isArray(channelList)) {
      for (const channel of channelList) {
        const c = channel as Record<string, unknown>;
        const channelId = c.id as string | undefined;
        const channelName = c.displayName as string | undefined;
        
        if (!channelId || !channelName) continue;
        
        // Channel has groupId directly, and channelType as a number
        const groupId = (c.groupId as string) || '';
        // Map numeric channelType to string (0=Standard, 1=Private, 2=Shared)
        const channelTypeNum = c.channelType as number | undefined;
        const channelType = channelTypeNum === 1 ? 'Private' 
          : channelTypeNum === 2 ? 'Shared' 
          : 'Standard';
        
        channels.push({
          channelId,
          channelName,
          teamName: displayName,
          teamId: groupId,
          channelType,
          description: c.description as string | undefined,
          isMember: true, // User is always a member for channels returned by this API
        });
      }
    }
    
    results.push({
      teamId: threadId, // Use thread ID as team identifier
      teamName: displayName,
      threadId,
      description: t.description as string | undefined,
      channels,
    });
  }
  
  return results;
}

/**
 * Filters channels from the Teams List by name.
 * 
 * @param teams - Array of teams with channels from parseTeamsList
 * @param query - Search query (case-insensitive partial match)
 * @returns Matching channels flattened into a single array
 */
export function filterChannelsByName(
  teams: TeamWithChannels[],
  query: string
): ChannelSearchResult[] {
  const lowerQuery = query.toLowerCase();
  const results: ChannelSearchResult[] = [];
  
  for (const team of teams) {
    for (const channel of team.channels) {
      if (channel.channelName.toLowerCase().includes(lowerQuery)) {
        results.push(channel);
      }
    }
  }
  
  return results;
}
