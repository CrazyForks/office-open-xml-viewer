import { describe, it, expect } from 'vitest';
import {
  sanitizeHyperlinkUrl,
  hyperlinkUrlScheme,
  openExternalHyperlink,
  DEFAULT_ALLOWED_HYPERLINK_SCHEMES,
} from './hyperlink.js';

/**
 * IX1 hyperlink URL sanitiser. External hyperlink targets come straight out of
 * an untrusted document's relationship part, so before a click navigates
 * anywhere the scheme must be checked against an allowlist. The dangerous cases
 * (`javascript:`, `data:`, `vbscript:`, `file:`) must be blocked even when the
 * author obfuscates them with leading whitespace or embedded control characters
 * that browsers strip when resolving the URL.
 */
describe('sanitizeHyperlinkUrl (external link scheme allowlist)', () => {
  it('allows the safe web / contact schemes', () => {
    expect(sanitizeHyperlinkUrl('https://example.com/a?b=1#c')).toBe('https://example.com/a?b=1#c');
    expect(sanitizeHyperlinkUrl('http://example.com')).toBe('http://example.com');
    expect(sanitizeHyperlinkUrl('mailto:hi@example.com')).toBe('mailto:hi@example.com');
    expect(sanitizeHyperlinkUrl('tel:+15551234567')).toBe('tel:+15551234567');
  });

  it('blocks script / data / local-file schemes (returns null)', () => {
    expect(sanitizeHyperlinkUrl('javascript:alert(1)')).toBeNull();
    expect(sanitizeHyperlinkUrl('data:text/html,<script>alert(1)</script>')).toBeNull();
    expect(sanitizeHyperlinkUrl('vbscript:msgbox(1)')).toBeNull();
    expect(sanitizeHyperlinkUrl('file:///etc/passwd')).toBeNull();
    expect(sanitizeHyperlinkUrl('about:blank')).toBeNull();
    expect(sanitizeHyperlinkUrl('blob:https://x/y')).toBeNull();
  });

  it('is case-insensitive on the scheme', () => {
    expect(sanitizeHyperlinkUrl('HTTPS://example.com')).toBe('HTTPS://example.com');
    expect(sanitizeHyperlinkUrl('JavaScript:alert(1)')).toBeNull();
    expect(sanitizeHyperlinkUrl('JAVASCRIPT:alert(1)')).toBeNull();
  });

  it('sees through leading whitespace and embedded control chars', () => {
    // Browsers strip TAB/LF/CR and leading spaces when resolving a URL, so
    // "java\tscript:" IS the javascript scheme and must be blocked.
    const tab = String.fromCharCode(9);
    const lf = String.fromCharCode(10);
    const cr = String.fromCharCode(13);
    expect(sanitizeHyperlinkUrl(`  javascript:alert(1)`)).toBeNull();
    expect(sanitizeHyperlinkUrl(`java${tab}script:alert(1)`)).toBeNull();
    expect(sanitizeHyperlinkUrl(`java${lf}script:alert(1)`)).toBeNull();
    expect(sanitizeHyperlinkUrl(`${cr}${tab} javascript:alert(1)`)).toBeNull();
    // A leading NUL byte likewise must not hide the scheme.
    expect(sanitizeHyperlinkUrl(`${String.fromCharCode(0)}javascript:alert(1)`)).toBeNull();
  });

  it('treats a scheme-less (relative / anchor / protocol-relative) target as safe', () => {
    expect(sanitizeHyperlinkUrl('page.html')).toBe('page.html');
    expect(sanitizeHyperlinkUrl('#section')).toBe('#section');
    expect(sanitizeHyperlinkUrl('//example.com/x')).toBe('//example.com/x');
    expect(sanitizeHyperlinkUrl('/root/rel')).toBe('/root/rel');
  });

  it('blocks the empty string', () => {
    expect(sanitizeHyperlinkUrl('')).toBeNull();
  });

  it('honours a custom allowlist', () => {
    // An intranet viewer that opts into file: links.
    const allow = ['https', 'file'];
    expect(sanitizeHyperlinkUrl('file:///share/doc', allow)).toBe('file:///share/doc');
    // mailto is no longer allowed under the custom list.
    expect(sanitizeHyperlinkUrl('mailto:x@y.z', allow)).toBeNull();
  });

  it('default allowlist is exactly the safe four', () => {
    expect([...DEFAULT_ALLOWED_HYPERLINK_SCHEMES]).toEqual(['http', 'https', 'mailto', 'tel']);
  });
});

describe('openExternalHyperlink (default new-tab open policy)', () => {
  it('opens an allowed URL in a new tab with noopener,noreferrer', () => {
    const calls: Array<[string, string, string]> = [];
    const win = { open: (u: string, t: string, f: string) => { calls.push([u, t, f]); return null; } };
    const ok = openExternalHyperlink('https://example.com', undefined, win);
    expect(ok).toBe(true);
    expect(calls).toEqual([['https://example.com', '_blank', 'noopener,noreferrer']]);
  });

  it('does NOT open a blocked scheme and returns false', () => {
    let opened = false;
    const win = { open: () => { opened = true; return null; } };
    const ok = openExternalHyperlink('javascript:alert(1)', undefined, win);
    expect(ok).toBe(false);
    expect(opened).toBe(false);
  });

  it('returns false (no throw) when there is no window', () => {
    expect(openExternalHyperlink('https://example.com', undefined, undefined)).toBe(false);
  });
});

describe('hyperlinkUrlScheme', () => {
  it('extracts and lowercases the scheme', () => {
    expect(hyperlinkUrlScheme('HTTPS://x')).toBe('https');
    expect(hyperlinkUrlScheme('mailto:a@b')).toBe('mailto');
    expect(hyperlinkUrlScheme('ms-word:ofe|u|x')).toBe('ms-word');
  });

  it('returns null when there is no scheme', () => {
    expect(hyperlinkUrlScheme('relative/path')).toBeNull();
    expect(hyperlinkUrlScheme('//host/path')).toBeNull();
    expect(hyperlinkUrlScheme(':leadingcolon')).toBeNull();
    expect(hyperlinkUrlScheme('')).toBeNull();
  });
});
