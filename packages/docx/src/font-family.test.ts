import { describe, it, expect } from 'vitest';
import { normalizeFontFamily } from './renderer.js';

describe('normalizeFontFamily — Arabic substitute fonts', () => {
  it('puts the Arabic substitute first so Latin/digits resolve from the same family as Arabic', () => {
    // Sakkal Majalla is family="auto" in fontTable; the run carries both
    // Arabic glyphs and Latin/digits. The Arabic substitute must lead the chain
    // (before any CJK sans face) so Latin/digits don't leak to Noto Sans JP.
    const chain = normalizeFontFamily('Sakkal Majalla');
    expect(chain.startsWith('"Sakkal Majalla", "Noto Naskh Arabic"')).toBe(true);
    // No CJK sans face before the Arabic substitute.
    const naskhIdx = chain.indexOf('Noto Naskh Arabic');
    const cjkIdx = chain.indexOf('Noto Sans JP');
    expect(naskhIdx).toBeGreaterThan(0);
    expect(naskhIdx).toBeLessThan(cjkIdx);
  });

  it('routes traditional Naskh faces to a serif Latin companion', () => {
    // Word's PDF export of sample-7 renders Sakkal Majalla's Latin with serifs,
    // so a serif Latin generic precedes the sans generics.
    const chain = normalizeFontFamily('Traditional Arabic');
    expect(chain).toContain('"Noto Serif"');
    expect(chain.endsWith('serif')).toBe(true);
    expect(chain.indexOf('Noto Serif')).toBeLessThan(chain.indexOf('Noto Sans JP'));
  });

  it('keys off the family, not a hardcoded string — case-insensitive', () => {
    expect(normalizeFontFamily('sakkal majalla').startsWith('"sakkal majalla", "Noto Naskh Arabic"')).toBe(true);
  });

  it('routes geometric Arabic faces (Univers Next Arabic) to a sans chain', () => {
    const chain = normalizeFontFamily('Univers Next Arabic');
    expect(chain.startsWith('"Univers Next Arabic", "Noto Sans Arabic"')).toBe(true);
    expect(chain.endsWith('sans-serif')).toBe(true);
  });
});

describe('normalizeFontFamily — Latin defaults keep JP companion + add non-CJK scripts', () => {
  it('keeps Noto Sans JP as the East-Asian companion for a Latin sans font', () => {
    const chain = normalizeFontFamily('Arial');
    expect(chain.startsWith('"Arial", "Noto Sans JP"')).toBe(true);
    expect(chain.endsWith('sans-serif')).toBe(true);
    // Arabic + non-CJK script Notos present so glyphs degrade to real web fonts.
    expect(chain).toContain('"Noto Naskh Arabic"');
    expect(chain).toContain('"Noto Sans Hebrew"');
    expect(chain).toContain('"Noto Sans Thai"');
    expect(chain).toContain('"Noto Sans Devanagari"');
  });

  it('keeps the serif JP companion for a Latin serif font', () => {
    const chain = normalizeFontFamily('Times New Roman');
    expect(chain.startsWith('"Times New Roman", "Yu Mincho"')).toBe(true);
    expect(chain).toContain('"Noto Serif JP"');
    expect(chain).toContain('"Noto Serif Hebrew"');
    expect(chain.endsWith('serif')).toBe(true);
  });
});

describe('normalizeFontFamily — CJK language-specific Noto ordering', () => {
  it('puts Noto Sans KR first for Korean sans faces (Malgun Gothic, Gulim, Dotum, 돋움)', () => {
    for (const f of ['Malgun Gothic', 'Gulim', 'Dotum', '돋움']) {
      const chain = normalizeFontFamily(f);
      expect(chain, `${f}`).toContain('"Noto Sans KR"');
      // KR must precede JP so shared Han renders with Korean shapes.
      expect(chain.indexOf('Noto Sans KR')).toBeLessThan(
        chain.indexOf('Noto Sans JP') === -1 ? Infinity : chain.indexOf('Noto Sans JP'),
      );
    }
  });

  it('routes Korean serif faces (Batang) to Noto Serif KR', () => {
    const chain = normalizeFontFamily('Batang');
    expect(chain).toContain('"Noto Serif KR"');
    expect(chain.endsWith('serif')).toBe(true);
    expect(chain.indexOf('Noto Serif KR')).toBeLessThan(chain.indexOf('Noto Serif JP'));
  });

  it('puts Noto Sans SC first for Simplified Chinese faces (SimSun→serif, YaHei→sans)', () => {
    // SimSun is a song (serif) face → Noto Serif SC.
    const simsun = normalizeFontFamily('SimSun');
    expect(simsun).toContain('"Noto Serif SC"');
    expect(simsun.indexOf('Noto Serif SC')).toBeLessThan(
      simsun.indexOf('Noto Serif JP') === -1 ? Infinity : simsun.indexOf('Noto Serif JP'),
    );
    // Microsoft YaHei is a sans face → Noto Sans SC.
    const yahei = normalizeFontFamily('Microsoft YaHei');
    expect(yahei).toContain('"Noto Sans SC"');
    expect(yahei.indexOf('Noto Sans SC')).toBeLessThan(
      yahei.indexOf('Noto Sans JP') === -1 ? Infinity : yahei.indexOf('Noto Sans JP'),
    );
  });

  it('puts Noto Sans TC first for Traditional Chinese faces (PMingLiU→serif, JhengHei→sans)', () => {
    const pming = normalizeFontFamily('PMingLiU');
    expect(pming).toContain('"Noto Serif TC"');
    const jheng = normalizeFontFamily('Microsoft JhengHei');
    expect(jheng).toContain('"Noto Sans TC"');
    expect(jheng.indexOf('Noto Sans TC')).toBeLessThan(
      jheng.indexOf('Noto Sans SC') === -1 ? Infinity : jheng.indexOf('Noto Sans SC'),
    );
  });

  it('keeps Japanese faces on Noto JP (regression — Yu Gothic, Meiryo, MS Mincho)', () => {
    expect(normalizeFontFamily('Yu Gothic')).toContain('"Noto Sans JP"');
    expect(normalizeFontFamily('Meiryo')).toContain('"Noto Sans JP"');
    // MS Mincho is serif → Noto Serif JP.
    expect(normalizeFontFamily('MS Mincho')).toContain('"Noto Serif JP"');
  });
});
