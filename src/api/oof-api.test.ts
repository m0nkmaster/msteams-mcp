import { expect, it, vi } from 'vitest';
import { ok } from '../types/result.js';

vi.mock('../utils/auth-guards.js', () => ({ requireGraphTokenAsync: vi.fn() }));
vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));

import { requireGraphTokenAsync } from '../utils/auth-guards.js';
import { httpRequest } from '../utils/http.js';
import { setAutomaticReplies } from './oof-api.js';

it('sends scheduled automatic replies with a UTC end and default external message', async () => {
  vi.mocked(requireGraphTokenAsync).mockResolvedValue(ok('graph-token'));
  vi.mocked(httpRequest).mockResolvedValue(ok({
    status: 200,
    headers: new Headers(),
    data: {},
  }));

  await setAutomaticReplies({
    enabled: true,
    internalMessage: 'Out of office',
    until: '2026-09-17T12:30:00+02:00',
  });

  const request = vi.mocked(httpRequest).mock.calls[0]?.[1];
  const setting = JSON.parse(String(request?.body)).automaticRepliesSetting;
  expect(request?.method).toBe('PATCH');
  expect(setting.status).toBe('scheduled');
  expect(setting.scheduledEndDateTime).toEqual({
    dateTime: '2026-09-17T10:30:00',
    timeZone: 'UTC',
  });
  expect(setting.externalReplyMessage).toBe('Out of office');
});
