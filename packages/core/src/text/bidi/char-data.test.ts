import { describe, it, expect } from 'vitest';
import { bidiClass, mirror, bracket, UNICODE_VERSION } from './char-data.js';

describe('bidiClass', () => {
  it('classifies strong types', () => {
    expect(bidiClass(0x41)).toBe('L'); // LATIN CAPITAL LETTER A
    expect(bidiClass(0x5d0)).toBe('R'); // HEBREW LETTER ALEF
    expect(bidiClass(0x627)).toBe('AL'); // ARABIC LETTER ALEF
  });
  it('classifies numbers and separators', () => {
    expect(bidiClass(0x30)).toBe('EN'); // DIGIT ZERO
    expect(bidiClass(0x660)).toBe('AN'); // ARABIC-INDIC DIGIT ZERO
    expect(bidiClass(0x20)).toBe('WS'); // SPACE
    expect(bidiClass(0x28)).toBe('ON'); // LEFT PARENTHESIS
  });
  it('honors @missing block defaults for unassigned code points', () => {
    // U+05EB is unassigned but in the Hebrew block -> default R.
    expect(bidiClass(0x5eb)).toBe('R');
  });
});

describe('mirror', () => {
  it('returns the mirror glyph for mirrored characters', () => {
    expect(mirror(0x28)).toBe(0x29);
    expect(mirror(0x29)).toBe(0x28);
    expect(mirror(0x3c)).toBe(0x3e); // < -> >
  });
  it('returns null for non-mirrored characters', () => {
    expect(mirror(0x41)).toBeNull();
    expect(mirror(0x627)).toBeNull();
  });
});

describe('bracket', () => {
  it('returns paired bracket info', () => {
    expect(bracket(0x28)).toEqual({ pair: 0x29, type: 'o' });
    expect(bracket(0x29)).toEqual({ pair: 0x28, type: 'c' });
    expect(bracket(0x5b)).toEqual({ pair: 0x5d, type: 'o' });
  });
  it('returns null for non-brackets', () => {
    expect(bracket(0x41)).toBeNull();
    expect(bracket(0x3c)).toBeNull(); // < is mirrored but not a paired bracket
  });
});

describe('metadata', () => {
  it('records the Unicode version', () => {
    expect(UNICODE_VERSION).toMatch(/^\d+\.\d+\.\d+$/);
  });
});
