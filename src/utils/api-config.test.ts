/**
 * Unit tests for api-config utilities.
 */

import { describe, it, expect } from 'vitest';
import {
  DEFAULT_TEAMS_BASE_URL,
  DEFAULT_SUBSTRATE_BASE_URL,
  SUBSTRATE_API,
  CHATSVC_API,
  CALENDAR_API,
  CSA_API,
  getTeamsHeaders,
  getBearerHeaders,
  getSkypeAuthHeaders,
  getCsaHeaders,
  getMessagingHeaders,
} from './api-config.js';

describe('DEFAULT_TEAMS_BASE_URL', () => {
  it('has a valid commercial teams URL', () => {
    expect(DEFAULT_TEAMS_BASE_URL).toBe('https://teams.microsoft.com');
  });
});

describe('DEFAULT_SUBSTRATE_BASE_URL', () => {
  it('has a valid commercial substrate URL', () => {
    expect(DEFAULT_SUBSTRATE_BASE_URL).toBe('https://substrate.office.com');
  });
});

describe('SUBSTRATE_API', () => {
  it('has search endpoint', () => {
    expect(SUBSTRATE_API.search).toContain('substrate.office.com');
    expect(SUBSTRATE_API.search).toContain('query');
  });

  it('has suggestions endpoint', () => {
    expect(SUBSTRATE_API.suggestions).toContain('substrate.office.com');
    expect(SUBSTRATE_API.suggestions).toContain('suggestions');
  });

  it('has frequent contacts endpoint', () => {
    expect(SUBSTRATE_API.frequentContacts).toContain('peoplecache');
  });

  it('has people search endpoint', () => {
    expect(SUBSTRATE_API.peopleSearch).toContain('powerbar');
  });

  it('has channel search endpoint', () => {
    expect(SUBSTRATE_API.channelSearch).toContain('TeamsChannel');
  });
});

describe('CHATSVC_API', () => {
  it('builds messages URL', () => {
    const url = CHATSVC_API.messages('amer-02', '19:abc@thread.tacv2');
    expect(url).toContain('chatsvc');
    expect(url).toContain('amer-02');
    expect(url).toContain('19%3Aabc%40thread.tacv2');
  });

  it('builds messages URL with replyToMessageId', () => {
    const url = CHATSVC_API.messages('amer-02', '19:abc@thread.tacv2', '1705760000000');
    // URL is encoded: %3B is semicolon, %3D is equals
    expect(url).toContain('%3Bmessageid%3D1705760000000');
  });

  it('builds conversation URL', () => {
    const url = CHATSVC_API.conversation('amer-02', '19:abc@thread.tacv2');
    expect(url).toContain('conversations');
    expect(url).toContain('19%3Aabc%40thread.tacv2');
  });

  it('builds messageMetadata URL', () => {
    const url = CHATSVC_API.messageMetadata('amer-02', '19:abc@thread.tacv2', 'msg123');
    expect(url).toContain('rcmetadata');
    expect(url).toContain('msg123');
  });

  it('builds editMessage URL', () => {
    const url = CHATSVC_API.editMessage('amer-02', '19:abc@thread.tacv2', 'msg123');
    expect(url).toContain('messages');
    expect(url).toContain('msg123');
  });

  it('builds deleteMessage URL', () => {
    const url = CHATSVC_API.deleteMessage('amer-02', '19:abc@thread.tacv2', 'msg123');
    expect(url).toContain('softDelete');
  });

  it('builds consumptionHorizons URL', () => {
    const url = CHATSVC_API.consumptionHorizons('amer-02', '19:abc@thread.tacv2');
    expect(url).toContain('consumptionhorizons');
  });

  it('builds updateConsumptionHorizon URL', () => {
    const url = CHATSVC_API.updateConsumptionHorizon('amer-02', '19:abc@thread.tacv2');
    expect(url).toContain('consumptionhorizon');
  });

  it('builds activityFeed URL', () => {
    const url = CHATSVC_API.activityFeed('amer-02');
    expect(url).toContain('conversations');
  });

  it('builds messageEmotions URL', () => {
    const url = CHATSVC_API.messageEmotions('amer-02', '19:abc@thread.tacv2', 'msg123');
    expect(url).toContain('emotions');
  });

  it('builds createThread URL', () => {
    const url = CHATSVC_API.createThread('amer-02');
    expect(url).toContain('threads');
  });

  it('builds singleMessage URL', () => {
    const url = CHATSVC_API.singleMessage('amer-02', '19:abc@thread.tacv2', 'msg123');
    expect(url).toContain('msg123');
  });

  it('accepts custom baseUrl for GCC support', () => {
    const url = CHATSVC_API.messages('amer-02', '19:abc@thread.tacv2', undefined, 'https://teams.microsoft.us');
    expect(url).toContain('teams.microsoft.us');
  });
});

describe('CALENDAR_API', () => {
  it('builds calendarView URL with partition', () => {
    const url = CALENDAR_API.calendarView('amer-02', true);
    expect(url).toContain('/mt/part/');
    expect(url).toContain('amer-02');
    expect(url).toContain('calendarView');
  });

  it('builds calendarView URL without partition', () => {
    const url = CALENDAR_API.calendarView('emea', false);
    expect(url).toContain('/mt/emea');
    expect(url).toContain('calendarView');
  });

  it('accepts custom baseUrl', () => {
    const url = CALENDAR_API.calendarView('amer-02', true, 'https://teams.microsoft.us');
    expect(url).toContain('teams.microsoft.us');
  });
});

describe('CSA_API', () => {
  it('builds conversationFolders URL', () => {
    const url = CSA_API.conversationFolders('amer-02');
    expect(url).toContain('csa');
    expect(url).toContain('conversationFolders');
  });

  it('builds teamsList URL', () => {
    const url = CSA_API.teamsList('amer-02');
    expect(url).toContain('teams');
    expect(url).toContain('api/v3');
  });

  it('builds customEmojis URL', () => {
    const url = CSA_API.customEmojis('amer-02');
    expect(url).toContain('customemoji');
    expect(url).toContain('metadata');
  });

  it('accepts custom baseUrl for government clouds', () => {
    const url = CSA_API.teamsList('amer-02', 'https://teams.microsoft.us');
    expect(url).toContain('teams.microsoft.us');
  });
});

describe('getTeamsHeaders', () => {
  it('returns headers with content-type and origin', () => {
    const headers = getTeamsHeaders();
    
    expect(headers['Content-Type']).toBe('application/json');
    expect(headers['Accept']).toBe('application/json');
    expect(headers['Origin']).toBe(DEFAULT_TEAMS_BASE_URL);
    expect(headers['Referer']).toBe(`${DEFAULT_TEAMS_BASE_URL}/`);
  });

  it('uses custom baseUrl for origin and referer', () => {
    const headers = getTeamsHeaders('https://teams.microsoft.us');
    
    expect(headers['Origin']).toBe('https://teams.microsoft.us');
    expect(headers['Referer']).toBe('https://teams.microsoft.us/');
  });
});

describe('getBearerHeaders', () => {
  it('adds Authorization header', () => {
    const headers = getBearerHeaders('my-token');
    
    expect(headers['Authorization']).toBe('Bearer my-token');
    expect(headers['Content-Type']).toBe('application/json');
  });

  it('uses custom baseUrl', () => {
    const headers = getBearerHeaders('my-token', 'https://teams.microsoft.us');
    
    expect(headers['Origin']).toBe('https://teams.microsoft.us');
  });
});

describe('getSkypeAuthHeaders', () => {
  it('adds both skypetoken and Authorization headers', () => {
    const headers = getSkypeAuthHeaders('skype-token', 'bearer-token');
    
    expect(headers['Authentication']).toBe('skypetoken=skype-token');
    expect(headers['Authorization']).toBe('Bearer bearer-token');
  });
});

describe('getCsaHeaders', () => {
  it('adds skypetoken and CSA bearer headers', () => {
    const headers = getCsaHeaders('skype-token', 'csa-token');
    
    expect(headers['Authentication']).toBe('skypetoken=skype-token');
    expect(headers['Authorization']).toBe('Bearer csa-token');
  });
});

describe('getMessagingHeaders', () => {
  it('includes client version header', () => {
    const headers = getMessagingHeaders('skype-token', 'bearer-token');
    
    expect(headers['X-Ms-Client-Version']).toBeDefined();
    expect(headers['X-Ms-Client-Version']).toMatch(/^\d+\/\d+\.\d+\.\d+/);
  });

  it('includes skypetoken and authorization', () => {
    const headers = getMessagingHeaders('skype-token', 'bearer-token');
    
    expect(headers['Authentication']).toContain('skypetoken');
    expect(headers['Authorization']).toContain('Bearer');
  });
});
