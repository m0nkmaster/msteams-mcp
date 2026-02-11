/**
 * Transcript API client for meeting transcript operations.
 * 
 * Uses Substrate WorkingSetFiles API to fetch meeting transcripts.
 * The transcript is embedded as JSON in the WorkingSetFiles response,
 * using the same Substrate token already used for search.
 * 
 * Flow: threadId → Substrate WorkingSetFiles (filter by MeetingThreadId) → parse TranscriptJson
 */

import { httpRequest } from '../utils/http.js';
import { type Result, ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import { requireSubstrateTokenAsync } from '../utils/auth-guards.js';
import { formatTranscriptText, type TranscriptEntry } from '../utils/parsers.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/** Result of fetching a transcript. */
export interface TranscriptResult {
  /** Meeting title from the recording metadata. */
  meetingTitle?: string;
  /** Meeting thread ID used for the lookup. */
  threadId: string;
  /** When the recording started. */
  recordingStartTime?: string;
  /** When the recording ended. */
  recordingEndTime?: string;
  /** Parsed transcript entries with timestamps and speakers. */
  entries: TranscriptEntry[];
  /** Formatted readable transcript text. */
  formattedText: string;
  /** Number of transcript entries. */
  entryCount: number;
  /** List of unique speakers in the transcript. */
  speakers: string[];
}

// ─────────────────────────────────────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────────────────────────────────────

/** Substrate WorkingSetFiles endpoint for finding meeting recordings + transcripts. */
const WORKING_SET_FILES_URL = 'https://substrate.office.com/api/beta/me/WorkingSetFiles/';

/** Fields to select from WorkingSetFiles. */
const WORKING_SET_SELECT = [
  'SharePointItem',
  'Visualization',
  'ItemProperties/Default/MeetingCallId',
  'ItemProperties/Default/DriveId',
  'ItemProperties/Default/RecordingStartDateTime',
  'ItemProperties/Default/RecordingEndDateTime',
  'ItemProperties/Default/TranscriptJson',
  'ItemProperties/Default/DocumentLink',
].join(',');

// ─────────────────────────────────────────────────────────────────────────────
// API Functions
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Gets the transcript for a meeting by its thread ID.
 * 
 * Uses Substrate WorkingSetFiles to find the recording associated with the
 * meeting thread, then extracts the embedded TranscriptJson.
 * 
 * @param threadId - The meeting thread ID (e.g., "19:meeting_xxx@thread.v2")
 * @param meetingDate - ISO date string of the meeting (used to narrow the search window)
 * @returns Parsed transcript with entries and formatted text
 */
export async function getTranscriptContent(
  threadId: string,
  meetingDate?: string
): Promise<Result<TranscriptResult>> {
  const authResult = await requireSubstrateTokenAsync();
  if (!authResult.ok) return authResult;
  const token = authResult.value;

  // Build date filter: search ±1 day around the meeting date, or last 30 days
  let dateFilter = '';
  if (meetingDate) {
    const date = new Date(meetingDate);
    const dayBefore = new Date(date);
    dayBefore.setDate(dayBefore.getDate() - 1);
    const dayAfter = new Date(date);
    dayAfter.setDate(dayAfter.getDate() + 1);
    dateFilter = ` AND FileCreatedTime gt ${dayBefore.toISOString()} AND FileCreatedTime lt ${dayAfter.toISOString()}`;
  }

  const filter = `ItemProperties/Default/MeetingThreadId eq '${threadId}'${dateFilter}`;
  const url = `${WORKING_SET_FILES_URL}?$filter=${encodeURIComponent(filter)}&$orderby=${encodeURIComponent('FileCreatedTime desc')}&$select=${encodeURIComponent(WORKING_SET_SELECT)}`;

  const response = await httpRequest<Record<string, unknown>>(
    url,
    {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json',
        'Content-Type': 'application/json',
        'Prefer': 'substrate.flexibleschema,outlook.data-source="Substrate",exchange.behavior="SubstrateFiles"',
      },
    }
  );

  if (!response.ok) {
    return response;
  }

  const data = response.value.data;
  const items = data.value as Array<Record<string, unknown>> | undefined;

  if (!items || items.length === 0) {
    return err(createError(
      ErrorCode.NOT_FOUND,
      'No recording found for this meeting. Transcription may not have been enabled, or the recording has not finished processing.',
      { suggestions: [
        'Check that transcription/recording was enabled during the meeting',
        'Wait a few minutes if the meeting just ended',
        'The meeting organiser must have recording enabled',
      ] }
    ));
  }

  const item = items[0];
  const props = (item.ItemProperties as Record<string, unknown>)?.Default as Record<string, unknown> | undefined;
  const viz = item.Visualization as Record<string, unknown> | undefined;

  // Extract TranscriptJson
  const transcriptJsonStr = props?.TranscriptJson as string | undefined;
  if (!transcriptJsonStr) {
    return err(createError(
      ErrorCode.NOT_FOUND,
      'Recording found but no transcript available. Transcription may not have been enabled for this meeting.',
      { suggestions: ['Check that transcription was enabled during the meeting'] }
    ));
  }

  // Parse the transcript JSON
  let transcriptData: { entries?: Array<Record<string, unknown>> };
  try {
    transcriptData = JSON.parse(transcriptJsonStr);
  } catch (parseError) {
    return err(createError(
      ErrorCode.UNKNOWN,
      `Failed to parse transcript data: ${parseError instanceof Error ? parseError.message : 'unknown error'}`,
    ));
  }

  const rawEntries = transcriptData.entries || [];
  if (rawEntries.length === 0) {
    return err(createError(
      ErrorCode.NOT_FOUND,
      'Transcript is empty — no speech was detected during the meeting.',
    ));
  }

  // Map to TranscriptEntry format
  // Substrate returns timestamps with microsecond precision (extra 4 trailing zeros) — normalise to milliseconds
  const entries: TranscriptEntry[] = rawEntries.map(e => ({
    startTime: (e.startOffset as string || '').replace(/0{4}$/, ''),
    endTime: (e.endOffset as string || '').replace(/0{4}$/, ''),
    speaker: e.speakerDisplayName as string || '',
    text: e.text as string || '',
  }));

  const formattedText = formatTranscriptText(entries);
  const speakers = [...new Set(entries.map(e => e.speaker).filter(s => s.length > 0))];

  return ok({
    meetingTitle: viz?.Title as string | undefined,
    threadId,
    recordingStartTime: props?.RecordingStartDateTime as string | undefined,
    recordingEndTime: props?.RecordingEndDateTime as string | undefined,
    entries,
    formattedText,
    entryCount: entries.length,
    speakers,
  });
}
