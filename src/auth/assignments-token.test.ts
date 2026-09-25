import { beforeEach, describe, expect, it, vi } from 'vitest';
import { extractAssignmentsToken, invalidateAssignmentsToken } from './token-extractor.js';
import { readSessionState, writeSessionState, type SessionState } from './session-store.js';
import { ASSIGNMENTS_APP_ID } from '../constants.js';

vi.mock('./session-store.js', async importOriginal => ({
  ...await importOriginal<typeof import('./session-store.js')>(),
  readSessionState: vi.fn(), writeSessionState: vi.fn(),
}));

function token(aud: string, seconds = 3600) {
  return `eyJhbGciOiJub25lIn0.${Buffer.from(JSON.stringify({ aud, exp: Date.now() / 1000 + seconds })).toString('base64url')}.signature`;
}
function entry(secret: string, credentialType = 'AccessToken') {
  return { name: secret, value: JSON.stringify({ credentialType, target: 'EduAssignments.ReadWrite', secret }) };
}
function state(entries: ReturnType<typeof entry>[]): SessionState {
  return { cookies: [], origins: [{ origin: 'https://teams.microsoft.com', localStorage: entries }] };
}

beforeEach(() => vi.clearAllMocks());
describe('Assignments token selection', () => {
  it('rejects Graph tokens even when they have matching scopes and later expiry', () => {
    const expected = token(ASSIGNMENTS_APP_ID);
    const graph = token('00000003-0000-0000-c000-000000000000', 7200);
    expect(extractAssignmentsToken(state([entry(graph), entry(expected)]))?.token).toBe(expected);
    expect(extractAssignmentsToken(state([entry(graph)]))).toBeNull();
  });
  it('rejects non-access credentials and expired tokens', () => {
    expect(extractAssignmentsToken(state([entry(token(ASSIGNMENTS_APP_ID), 'IdToken'), entry(token(ASSIGNMENTS_APP_ID, -1))]))).toBeNull();
  });
  it('invalidates only the rejected token', () => {
    const rejected = token(ASSIGNMENTS_APP_ID);
    const unrelated = entry(token('graph'));
    vi.mocked(readSessionState).mockReturnValue(state([entry(rejected), unrelated]));
    invalidateAssignmentsToken(rejected);
    expect(vi.mocked(writeSessionState).mock.calls[0][0].origins[0].localStorage).toEqual([unrelated]);
  });
});
