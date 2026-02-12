/**
 * Typed interfaces for raw Teams API responses.
 * 
 * These represent the shape of data returned by Teams/Substrate APIs
 * before parsing. Using these instead of `Record<string, unknown>` 
 * provides better type safety and IDE autocompletion.
 */

// ─────────────────────────────────────────────────────────────────────────────
// Chatsvc Message API
// ─────────────────────────────────────────────────────────────────────────────

/** Raw message from chatsvc /messages endpoint. */
export interface RawChatsvcMessage {
  [key: string]: unknown;
  id?: string;
  messagetype?: string;
  content?: string;
  contenttype?: string;
  from?: string;
  imdisplayname?: string;
  displayName?: string;
  originalarrivaltime?: string;
  composetime?: string;
  clientmessageid?: string;
  conversationid?: string;
  conversationId?: string;
  /** Thread root message ID (for channel thread replies). */
  rootMessageId?: string;
  /** Message properties (deletetime, emotions, etc.). */
  properties?: {
    deletetime?: string;
    [key: string]: unknown;
  };
  /** Source conversation ID for activity items. */
  clumpId?: string;
  /** Reference ID for virtual conversations (saved/followed). */
  secondaryReferenceId?: string;
  /** Thread topic name. */
  threadtopic?: string;
  topic?: string;
  /** Parent message ID (for activity items). */
  parentMessageId?: string;
}

/** Raw response from chatsvc /messages endpoint. */
export interface RawMessagesResponse {
  messages?: RawChatsvcMessage[];
  syncState?: string;
}

/** Raw response from chatsvc /conversations/{id} endpoint. */
export interface RawConversationResponse {
  threadProperties?: {
    topicThreadTopic?: string;
    topic?: string;
    spaceThreadTopic?: string;
    threadtopic?: string;
    productThreadType?: string;
    groupId?: string;
    [key: string]: unknown;
  };
  members?: Array<{
    mri?: string;
    id?: string;
    friendlyName?: string;
    displayName?: string;
    name?: string;
    [key: string]: unknown;
  }>;
  [key: string]: unknown;
}

/** Raw response from chatsvc /threads endpoint (create group chat). */
export interface RawCreateThreadResponse {
  threadResource?: {
    id?: string;
    [key: string]: unknown;
  };
  id?: string;
  threadId?: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// Chatsvc Consumption Horizon API
// ─────────────────────────────────────────────────────────────────────────────

/** Raw response from chatsvc /consumptionhorizons endpoint. */
export interface RawConsumptionHorizonsResponse {
  id?: string;
  version?: string;
  consumptionhorizons?: Array<{
    id: string;
    consumptionhorizon: string;
  }>;
}

// ─────────────────────────────────────────────────────────────────────────────
// Substrate Search API
// ─────────────────────────────────────────────────────────────────────────────

/** Raw response from Substrate v2/query endpoint. */
export interface RawSubstrateSearchResponse {
  EntitySets?: Array<{
    ResultSets?: Array<{
      Total?: number;
      Results?: unknown[];
    }>;
  }>;
  [key: string]: unknown;
}

/** Raw response from Substrate suggestions endpoint (people/channels). */
export interface RawSubstrateSuggestionsResponse {
  Groups?: unknown[];
  [key: string]: unknown;
}

// ─────────────────────────────────────────────────────────────────────────────
// Files API
// ─────────────────────────────────────────────────────────────────────────────

/** Raw file item from Substrate AllFiles API. */
export interface RawFileItem {
  ItemType?: string;
  FileData?: {
    FileUrl?: string;
    WebUrl?: string;
    PreviewUrl?: string;
    FileName?: string;
    FileExtension?: string;
    SizeInBytes?: number;
    [key: string]: unknown;
  };
  WeblinkData?: {
    WebUrl?: string;
    Title?: string;
    Description?: string;
    [key: string]: unknown;
  };
  DateTimeLastModified?: string;
  SharedByDisplayName?: string;
  SharedBySmtp?: string;
  SharedDateTime?: string;
  SharedTime?: string;
  From?: {
    EmailAddress?: {
      Name?: string;
      Address?: string;
    };
  };
  [key: string]: unknown;
}

/** Raw response from Substrate AllFiles API. */
export interface RawAllFilesResponse {
  Items?: RawFileItem[];
  SkipToken?: string;
  [key: string]: unknown;
}

// ─────────────────────────────────────────────────────────────────────────────
// Transcript API
// ─────────────────────────────────────────────────────────────────────────────

/** Raw transcript item from Substrate WorkingSetFiles API. */
export interface RawTranscriptItem {
  ItemProperties?: {
    Default?: {
      TranscriptJson?: string;
      [key: string]: unknown;
    };
    [key: string]: unknown;
  };
  Visualization?: {
    Title?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
}

/** Raw response from Substrate WorkingSetFiles API. */
export interface RawWorkingSetFilesResponse {
  Items?: RawTranscriptItem[];
  [key: string]: unknown;
}
