import { describe, it, expect } from 'vitest';
import { layoutParagraph } from './renderer.js';
import type { Paragraph } from './types';
import type { TextRunData } from '@silurus/ooxml-core';

// PowerPoint-observed SEA (Thai/Lao/Khmer) line-fit behavior — issue #991,
// adjudicated against the PowerPoint PDF of the purpose-built adjudication
// fixture (12 placement boxes: a tab-led single-a:r δ sweep, δ = availEff −
// runW ∈ {2,10,25,45}pt, plus a 3-a:r variant and a no-tab variant at δ=10,
// each under thaiDist and left; over-full controls; trailing-space shrink
// sweep; record on the issue). PowerPoint's rules are NOT Word's (#1003):
//
// 1. GREEDY dictionary split. A no-space SEA chunk that fits a full line by
//    itself but not the remaining width of a non-empty line is SPLIT at the
//    last fitting dictionary boundary — a run prefix backfills line 1 in all
//    12 placement boxes. Word moves the chunk whole; PowerPoint never does.
//    ⚠️ Do NOT port the docx #1003 whole-chunk guard to pptx: the current
//    greedy fill IS the PowerPoint-verified behavior. (The fixture's only
//    engine divergence is lexicon granularity — PowerPoint's Thai dictionary
//    keeps one compound unbreakable where ICU splits it — not policy.)
//
// 2. Placement is independent of alignment (thaiDist == left in all 6 pairs),
//    of the a:r segmentation (3-way word-boundary split == single run), and
//    of a leading tab. (No tab gate here: pptx's tab-cell layout keeps every
//    tab-led line unwrapped until paint — `tabSeen` in layoutParagraph — so
//    tab-led text never reaches the SEA fill. That PowerPoint wraps tab-led
//    lines at all is a separate, script-independent divergence tracked apart
//    from the SEA rules pinned in this file.)
//
// 3. ZERO trailing-space shrink, same as Word: with algn="just", a paragraph-
//    final word overflowing by X wraps for every X > 0 (admits only X < 0),
//    and left renders identical line counts. pptx's exact-width fit (no
//    shrink budget) is already correct — do not add one for SEA.
//
// The mock advances 10px per UTF-16 unit, making the arithmetic explicit
// (all vocabulary is BMP). Thai combining marks count as full advances here —
// irrelevant, as every expected number below is computed in the same model.
// ICU boundaries (same segmenter as the engine): 'น้อยน้อยน้อยน้อย' = 16 units,
// boundaries at 4/8/12 (40px words); 'ดังนี้' = 6 units, no interior boundary;
// lead 'มาก ' = 4 units = 40px.

function mockCtx() {
  let font = '';
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    measureText: (s: string) => ({ width: s.length * 10 }),
    fillRect() {}, fillText() {},
    fillStyle: '', strokeStyle: '',
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

function run(text: string, over: Partial<TextRunData> = {}): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: 20, color: '000000', fontFamily: 'Arial', ...over,
  };
}

function para(runs: TextRunData[], alignment: string = 'l'): Paragraph {
  return {
    alignment, marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null, defBold: null,
    defItalic: null, defFontFamily: null, tabStops: [], eaLnBrk: true, runs,
  } as Paragraph;
}

const lineText = (line: { segments: { text: string }[] }): string =>
  line.segments.map((s) => s.text).join('');

const layoutLines = (p: Paragraph, width: number): string[] =>
  layoutParagraph(mockCtx(), p, width, 20, '000000', 1, 0).map(lineText);

const norm = (l: string): string => l.replace(/\s+/g, '');

// Lead token 'มาก ' = 40px. Chunk 'น้อยน้อยน้อยน้อย' = 160px, dictionary
// boundaries at 4/8/12 (40px words). Width 170: the chunk alone fits a full
// line (160 ≤ 170) but not the remainder (130) — Word moves it WHOLE (#1003);
// PowerPoint splits it at the last fitting boundary (12 units, 120 ≤ 130).
const THAI_CHUNK = 'มาก น้อยน้อยน้อยน้อย';

// 'มาก ' ×3 = 120px + final word 'ดังนี้' = 60px ⇒ natural end 180. The GT
// sweep admitted only X=−3 (width 183 here) and wrapped every X ≥ +1 (width
// 179): the shrink threshold is exactly zero.
const THAI_FINAL = 'มาก มาก มาก ดังนี้';

describe('issue #991 — pptx SEA fit (PowerPoint adjudication-fixture rules)', () => {
  it('Rule 1: a full-line-fitting no-space chunk is SPLIT greedily, not moved whole', () => {
    const lines = layoutLines(para([run(THAI_CHUNK)], 'thaiDist'), 170);
    expect(lines.length).toBe(2);
    // The docx/Word rule would leave only the lead word on line 1 — PowerPoint
    // backfills a dictionary prefix instead. This gate fails if #1003's
    // whole-chunk guard is ever ported to pptx.
    expect(norm(lines[0])).toBe('มากน้อยน้อยน้อย');
    expect(norm(lines[1])).toBe('น้อย');
  });

  it('Rule 2: placement is independent of alignment (left == thaiDist)', () => {
    const lines = layoutLines(para([run(THAI_CHUNK)]), 170);
    expect(lines.length).toBe(2);
    expect(norm(lines[0])).toBe('มากน้อยน้อยน้อย');
    expect(norm(lines[1])).toBe('น้อย');
  });

  it('Rule 2: placement is independent of the a:r segmentation (word-boundary splits)', () => {
    const lines = layoutLines(
      para([run('มาก '), run('น้อยน้อย'), run('น้อยน้อย')], 'thaiDist'),
      170,
    );
    expect(lines.length).toBe(2);
    expect(norm(lines[0])).toBe('มากน้อยน้อยน้อย');
    expect(norm(lines[1])).toBe('น้อย');
  });

  it('over-full control: a chunk wider than the line fills greedily at dictionary boundaries', () => {
    // Width 150 < chunk 160 ⇒ no line can hold the chunk whole; remaining
    // after the lead = 110 ⇒ split at 8 units (80 ≤ 110 < 120). PowerPoint's
    // forced breaks matched our boundaries in all 4 over-full boxes.
    const lines = layoutLines(para([run(THAI_CHUNK)], 'thaiDist'), 150);
    expect(lines.length).toBe(2);
    expect(norm(lines[0])).toBe('มากน้อยน้อย');
    expect(norm(lines[1])).toBe('น้อยน้อย');
  });

  it('control: a chunk that fits the remaining width stays on the line', () => {
    const lines = layoutLines(para([run('มาก น้อยน้อย')], 'thaiDist'), 170);
    expect(lines.length).toBe(1);
  });

  it('Rule 3: zero trailing-space shrink — a +1px overflowing final word wraps (thaiDist)', () => {
    const lines = layoutLines(para([run(THAI_FINAL)], 'thaiDist'), 179);
    expect(lines.length).toBe(2);
    expect(norm(lines[0])).toBe('มากมากมาก');
    expect(norm(lines[1])).toBe('ดังนี้');
  });

  it('Rule 3: same wrap under left alignment (fit is alignment-independent)', () => {
    const lines = layoutLines(para([run(THAI_FINAL)]), 179);
    expect(lines.length).toBe(2);
    expect(norm(lines[1])).toBe('ดังนี้');
  });

  it('Rule 3 control: X=−3 (the adjudicated admit case) stays on one line', () => {
    const lines = layoutLines(para([run(THAI_FINAL)], 'thaiDist'), 183);
    expect(lines.length).toBe(1);
  });
});
