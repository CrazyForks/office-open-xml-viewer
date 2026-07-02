import { describe, expect, it } from 'vitest';
import type { FontPreloadEntry } from '@silurus/ooxml-core';
import { PPTX_GOOGLE_FONTS } from './google-fonts';

// Verbatim snapshot of the PPTX Office-font substitute map BEFORE the shared
// registry consolidation (Phase 3 C7), excluding the SCRIPT_GOOGLE_FONTS spread
// (unchanged, shared already). Frozen as the oracle so the consolidated map's
// effective entries can only ADD keys, never drop or alter one.
const PPTX_GOOGLE_FONTS_OLD: Record<string, FontPreloadEntry> = {
  'calibri':           { url: 'https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Carlito' },
  'calibri light':     { url: 'https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Carlito' },
  'cambria':           { url: 'https://fonts.googleapis.com/css2?family=Caladea:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Caladea' },
  'cambria math':      { url: 'https://fonts.googleapis.com/css2?family=Caladea:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Caladea' },
  'nunito sans':       { url: 'https://fonts.googleapis.com/css2?family=Nunito+Sans:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'nunito':            { url: 'https://fonts.googleapis.com/css2?family=Nunito:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'open sans':         { url: 'https://fonts.googleapis.com/css2?family=Open+Sans:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'roboto':            { url: 'https://fonts.googleapis.com/css2?family=Roboto:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'lato':              { url: 'https://fonts.googleapis.com/css2?family=Lato:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'montserrat':        { url: 'https://fonts.googleapis.com/css2?family=Montserrat:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'poppins':           { url: 'https://fonts.googleapis.com/css2?family=Poppins:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'raleway':           { url: 'https://fonts.googleapis.com/css2?family=Raleway:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'playfair display':  { url: 'https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'sakkal majalla':      { url: 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Naskh Arabic' },
  'traditional arabic':  { url: 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Naskh Arabic' },
  'simplified arabic':   { url: 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Naskh Arabic' },
  'arabic typesetting':  { url: 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Naskh Arabic' },
  'univers next arabic': { url: 'https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Sans Arabic' },
  'noto naskh arabic':   { url: 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Naskh Arabic' },
  'noto sans arabic':    { url: 'https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Sans Arabic' },
};

describe('PPTX_GOOGLE_FONTS — shared registry consolidation (oracle)', () => {
  it('preserves every pre-consolidation entry byte-for-byte', () => {
    for (const [key, entry] of Object.entries(PPTX_GOOGLE_FONTS_OLD)) {
      expect(PPTX_GOOGLE_FONTS[key], `entry "${key}"`).toEqual(entry);
    }
  });

  it('adds only the safe, documented key Ubuntu', () => {
    // pptx already carried the full web-font + Office-substitute set. The shared
    // registry additionally contributes "ubuntu" (a generic Google web font, no
    // format affinity): a slide that requests Ubuntu now measures glyphs with
    // the real face instead of a narrower system sans. Purely additive.
    const oldKeys = new Set(Object.keys(PPTX_GOOGLE_FONTS_OLD));
    const added = Object.keys(PPTX_GOOGLE_FONTS).filter(
      (k) => !oldKeys.has(k) && !k.startsWith('noto '),
    );
    expect(new Set(added)).toEqual(new Set(['ubuntu']));
    expect(PPTX_GOOGLE_FONTS['ubuntu'].url).toMatch(/family=Ubuntu(?:[:&]|$)/);
    expect(PPTX_GOOGLE_FONTS['ubuntu'].loadFamily).toBeUndefined();
  });
});
