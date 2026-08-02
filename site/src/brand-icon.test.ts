import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const read = (path: string) => readFileSync(new URL(path, import.meta.url), 'utf8');
const nav = read('./components/Nav.astro');
const base = read('./layouts/Base.astro');
const footerSurfaces = [
  read('./layouts/FormatPage.astro'),
  read('./pages/index.astro'),
  read('./pages/try.astro'),
  read('./pages/deprecations.astro'),
  read('./pages/errors.astro'),
  read('./pages/announcements/index.astro'),
];

describe('official-site brand icon', () => {
  it('keeps the header typographic and uses the library icon in every footer', () => {
    expect(nav).not.toContain('<BrandIcon />');
    expect(nav).not.toContain('nav-mark');
    for (const source of footerSurfaces) {
      expect(source).toContain('<BrandIcon />');
      expect(source).not.toContain('nav-mark');
    }
  });

  it('uses the same PNG as the browser icon', () => {
    expect(base).toContain('<link rel="icon" type="image/png"');
    expect(base).toContain('icon.png');
    expect(base).not.toContain('favicon.svg');
  });
});
