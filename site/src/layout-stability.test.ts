import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const globalCss = readFileSync(new URL('./styles/global.css', import.meta.url), 'utf8');
const nav = readFileSync(new URL('./components/Nav.astro', import.meta.url), 'utf8');
const base = readFileSync(new URL('./layouts/Base.astro', import.meta.url), 'utf8');
const formatPage = readFileSync(new URL('./layouts/FormatPage.astro', import.meta.url), 'utf8');
const apiReference = readFileSync(new URL('./components/ApiReference.astro', import.meta.url), 'utf8');

function darkHexToken(name: string): [number, number, number] {
  const darkTheme = globalCss.match(/:root\[data-theme='dark'\]\s*\{([\s\S]*?)\n\}/)?.[1] ?? '';
  const hex = darkTheme.match(new RegExp(`--${name}:\\s*#([0-9a-f]{6})`, 'i'))?.[1];
  if (!hex) throw new Error(`Missing dark theme token: ${name}`);
  return [0, 2, 4].map((start) => Number.parseInt(hex.slice(start, start + 2), 16)) as [number, number, number];
}

function lightHexToken(name: string): [number, number, number] {
  const lightTheme = globalCss.match(/:root\s*\{([\s\S]*?)\n\}/)?.[1] ?? '';
  const hex = lightTheme.match(new RegExp(`--${name}:\\s*#([0-9a-f]{6})`, 'i'))?.[1];
  if (!hex) throw new Error(`Missing light theme token: ${name}`);
  return [0, 2, 4].map((start) => Number.parseInt(hex.slice(start, start + 2), 16)) as [number, number, number];
}

const largeSurfaceTokens = [
  'paper',
  'hero-bg',
  'bg-elev',
  'bg-elev-2',
  'panel',
  'surface-muted',
  'border',
  'preview-top',
  'preview-bottom',
  'code-bg',
];

describe('official-site layout stability', () => {
  it('reserves the root scrollbar gutter across route changes', () => {
    expect(globalCss).toMatch(/html\s*\{[^}]*scrollbar-gutter:\s*stable;/);
  });

  it('sizes the shared library icon consistently in the header and footer', () => {
    expect(globalCss).toMatch(/\.brand-icon\s*\{[^}]*width:\s*28px;[^}]*height:\s*28px;/);
  });

  it('renders the theme label from the pre-paint document theme', () => {
    expect(nav).toContain('<span class="theme-label" aria-hidden="true"></span>');
    expect(nav).not.toContain('>Theme</span>');
    expect(globalCss).toContain(".theme-label::before { content: 'Dark'; }");
    expect(globalCss).toContain(":root[data-theme='dark'] .theme-label::before { content: 'Light'; }");
    expect(base).not.toContain("querySelector<HTMLElement>('[data-theme-label]')");
  });

  it('keeps format detail labels typographic rather than adding competing colour dots', () => {
    expect(formatPage).toContain('<p class="eyebrow">{name}</p>');
    expect(formatPage).not.toContain('fp-dot');
    expect(formatPage).not.toContain('color: string');
  });

  it('uses the shared subtle shadow for DOCX pages and PPTX slides', () => {
    expect(globalCss).toMatch(/\.lv-page\s*\{[^}]*box-shadow:\s*var\(--document-shadow\);/);
    expect(globalCss).toMatch(/\.demo-page\s*\{[^}]*box-shadow:\s*var\(--document-shadow\);/);
  });

  it('keeps large dark-theme surfaces neutral so lime remains an accent', () => {
    for (const token of largeSurfaceTokens) {
      const channels = darkHexToken(token);
      expect(Math.max(...channels) - Math.min(...channels), token).toBeLessThanOrEqual(8);
    }
  });

  it('keeps large light-theme surfaces neutral so lime remains an accent', () => {
    for (const token of largeSurfaceTokens) {
      const channels = lightHexToken(token);
      expect(Math.max(...channels) - Math.min(...channels), token).toBeLessThanOrEqual(8);
    }
  });

  it('uses the stronger shared light accent throughout the API table without changing dark mode', () => {
    expect(apiReference).toContain('color: var(--accent-2)');
    expect(apiReference).toContain('color: var(--accent)');
    expect(lightHexToken('signal-ink')).toEqual([0x3b, 0x7e, 0x00]);
    expect(lightHexToken('accent-2')).toEqual([0x3b, 0x7e, 0x00]);
    expect(darkHexToken('signal-ink')).toEqual([0xc9, 0xff, 0x43]);
    expect(darkHexToken('accent-2')).toEqual([0xa9, 0xd3, 0x66]);
  });
});
