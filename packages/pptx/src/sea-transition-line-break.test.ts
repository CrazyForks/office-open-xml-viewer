import { describe, it, expect } from 'vitest';
import {
  DEFAULT_KINSOKU_RULES,
  seaMixedBreakOffsets,
  seaTransitionOffsets,
} from '@silurus/ooxml-core';
import { layoutParagraph } from './renderer.js';
import type { Paragraph } from './types';
import type { TextRunData } from '@silurus/ooxml-core';

// Issue #960 — pptx counterpart of the docx transition/mixed-run fix. A run that
// mixes SEA (Thai) with Latin/digit/CJK and has no spaces takes ONE token; the
// no-space script transitions must be break opportunities, and a mixed CJK+SEA
// token must break the CJK side per-character while keeping the Thai side at
// dictionary/cluster boundaries (adjudicated against Word/PowerPoint GT). char =
// 10px in the mock, so the fitter is deterministic; boundaries come from ICU.

function mockCtx() {
  let font = '';
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    measureText: (s: string) => ({ width: [...s].length * 10 }),
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
function para(runs: TextRunData[]): Paragraph {
  return {
    alignment: 'l', marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null, defBold: null,
    defItalic: null, defFontFamily: null, tabStops: [], eaLnBrk: true, runs,
  } as Paragraph;
}
const lineText = (line: { segments: { text: string }[] }): string =>
  line.segments.map((s) => s.text).join('');
function breakOffsets(texts: string[]): number[] {
  const out: number[] = [];
  let acc = 0;
  for (let i = 0; i < texts.length - 1; i++) { acc += texts[i].length; out.push(acc); }
  return out;
}
const legal = (t: string): Set<number> =>
  new Set(seaMixedBreakOffsets(t, { cjk: true, kinsoku: DEFAULT_KINSOKU_RULES }));

describe('pptx SEA↔non-SEA no-space transitions and mixed CJK+SEA (#960)', () => {
  const A2 = 'ราคาสินค้า1250บาทลดเหลือ990บาทประหยัด260บาทหรือ21เปอร์เซ็นต์ต่อชิ้น';
  const B = '日本語のテキストとภาษาไทยが同じ';

  it('a2 Thai↔digit: breaks at the digit seams, never mid-number', () => {
    const lines = layoutParagraph(mockCtx(), para([run(A2)]), 130, 20, '000000', 1, 0);
    const texts = lines.map(lineText);
    expect(texts.join('')).toBe(A2);
    expect(texts.length).toBeGreaterThan(1);
    const legalSet = legal(A2);
    for (const b of breakOffsets(texts)) expect(legalSet.has(b)).toBe(true);
    for (const b of breakOffsets(texts)) {
      const isDigit = (c: number) => c >= 0x30 && c <= 0x39;
      expect(isDigit(A2.codePointAt(b - 1)!) && isDigit(A2.codePointAt(b)!)).toBe(false);
    }
    const transitions = new Set(seaTransitionOffsets(A2));
    expect(breakOffsets(texts).some((b) => transitions.has(b))).toBe(true);
  });

  it('b CJK+Thai one run: each side keeps its rule — Thai span never torn', () => {
    const lines = layoutParagraph(mockCtx(), para([run(B)]), 100, 20, '000000', 1, 0);
    const texts = lines.map(lineText);
    expect(texts.join('')).toBe(B);
    expect(texts.length).toBeGreaterThan(1);
    const legalSet = legal(B);
    for (const b of breakOffsets(texts)) expect(legalSet.has(b)).toBe(true);
    // The contiguous Thai cluster run stays intact on ONE line.
    expect(texts.some((t) => t.includes('ภาษาไทย'))).toBe(true);
  });
});
