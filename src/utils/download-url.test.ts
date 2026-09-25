import { describe, expect, it } from 'vitest';
import { sharePointDownloadUrl } from './api-config.js';

describe('sharePointDownloadUrl', () => {
  it('escapes decoded paths once for OData, preserving hashes, percent signs and apostrophes', () => {
    const result = sharePointDownloadUrl("https://tenant.sharepoint.com/sites/demo/Shared%20Documents/O%27Brien%20%23100%25.md?web=1");
    expect(result.fileName).toBe("O'Brien #100%.md");
    expect(decodeURIComponent(new URL(result.url).pathname)).toBe("/sites/demo/_api/web/GetFileByServerRelativePath(decodedUrl='/sites/demo/Shared Documents/O''Brien #100%.md')/$value");
    expect(new URL(result.url).hash).toBe('');
  });
  it('accepts direct :r links', () => {
    expect(sharePointDownloadUrl('https://tenant.sharepoint.com/:w:/r/teams/demo/Documents/test.docx').fileName).toBe('test.docx');
  });
  it.each([
    'http://tenant.sharepoint.com/file',
    'https://tenant.sharepoint.com.evil.test/file',
    'https://tenant.sharepoint.com:8443/file',
    'https://user:password@tenant.sharepoint.com/file',
    'https://tenant.sharepoint.com/:w:/s/site/opaque',
    'https://tenant.sharepoint.com/sites/site/_layouts/15/Doc.aspx?id=1',
  ])('rejects unsupported URL %s', url => {
    expect(() => sharePointDownloadUrl(url)).toThrow();
  });
});
