import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
} from './types';

// Regression (mirror of cjk-glued-nonstarter-wrap.ts / sample-9 fb836d6, the
// OTHER glue direction): a Latin word ("…Roman") immediately followed by a CJK
// run that STARTS with a line-start-forbidden char ("、…") — the CJK run carries
// `joinPrev` (UAX#14 LB13 / §17.3.1.16, keeps "、" off a line head). The glued-
// group pre-flush summed the WHOLE breakable CJK run as one atomic unit glued to
// "Roman"; finding {Roman + run} too wide for the line, it flushed "Roman" to the
// next line ALONE even though "Roman" + "、" still fit. A `both`-justified line
// then held only "Times New" and stretched sparse (sample-16, "Times New Roman"
// paragraph).
//
// A breakable CJK follower is NOT atomic: only its LEADING run of non-starters
// (the part that would actually orphan at a line head per LB13) must stay glued
// to the Latin lead; the rest wraps on its own. So the pre-flush group width adds
// only that prefix's advance and stops. When "Roman" + "、" genuinely don't fit
// the pre-flush still fires (they start fresh together — no orphan, no mid-word
// Latin split); when there is room, "Roman" is placed and the CJK run splits
// normally, keeping "、" with "Roman".

const FONT_PX = 20; // glyph advance per glyph (Latin or CJK) in the stub (scale = 1)

function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  fillTextCalls: { text: string; x: number; y: number }[];
} {
  let font = `${FONT_PX}px serif`;
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const fillTextCalls: { text: string; x: number; y: number }[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText(text: string, x: number, y: number) { fillTextCalls.push({ text, x, y }); },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, fillTextCalls };
}

function textRun(text: string): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: FONT_PX, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  };
}

type DocRun = DocParagraph['runs'][number];

function gluedPara(texts: string[]): BodyElement {
  const p: DocParagraph = {
    alignment: 'both',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: texts.map((t) => ({ type: 'text', ...textRun(t) }) as DocRun),
    defaultFontSize: FONT_PX, defaultFontFamily: 'NotInMetrics',
    widowControl: false,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function section(overrides: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 320, pageHeight: 400, // contentWidth 320 → exactly 16 cells/line
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    ...overrides,
  };
}

function doc(body: BodyElement[], sec: SectionProps): DocxDocumentModel {
  return {
    section: sec, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

async function render(body: BodyElement[], sec: SectionProps) {
  const { canvas, fillTextCalls } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc(body, sec), canvas, 0, { dpr: 1, width: sec.pageWidth });
  return fillTextCalls;
}

function linesByY(calls: { text: string; x: number; y: number }[]) {
  const byY = new Map<number, { text: string; x: number }[]>();
  for (const c of calls) {
    const k = Math.round(c.y);
    (byY.get(k) ?? byY.set(k, []).get(k)!).push({ text: c.text, x: c.x });
  }
  return [...byY.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([, glyphs]) => glyphs.slice().sort((p, q) => p.x - q.x));
}

describe('Latin word + glued non-starter-led CJK run — keep "Roman" on its line', () => {
  it('does not orphan "Roman" alone behind a breakable CJK run that starts with "、"', async () => {
    // Layout (16 cells/line, every glyph 20px):
    //   line 1: 16 leading CJK cells (full).
    //   line 2: "Times " (6) + "New " (4) = 10 cells, 6 remaining.
    //           "Roman" (5) fits (1 cell left); the CJK run "、あいうえお" (6) is
    //           glued (joinPrev: starts with "、"). "Roman"+"、" = 6 cells = fits;
    //           "Roman"+whole run = 11 cells = does NOT fit.
    // Pre-fix: the pre-flush sums Roman+whole run (11 cells) → too wide → flushes
    //   "Roman" to line 3 alone, leaving line 2 = "Times New" (sparse when `both`).
    // Post-fix: the pre-flush adds only the leading "、" (1 cell) → Roman+"、" fits
    //   → "Roman" stays on line 2 and the CJK run splits to fill it.
    const calls = await render(
      [gluedPara(['ああああああああああああああああ', 'Times New Roman', '、あいうえお'])],
      section(),
    );
    expect(calls.length).toBeGreaterThan(0);

    const lines = linesByY(calls);
    // Find the line that holds the Latin word "Times".
    const latinLine = lines.find((g) => g.map((x) => x.text).join('').includes('Times'));
    expect(latinLine).toBeDefined();
    const latinText = latinLine!.map((g) => g.text).join('');

    // "Roman" must be on the SAME line as "Times"/"New" — NOT flushed down alone.
    // (Pre-fix this line was "Times New " only; "Roman" was orphaned below.)
    expect(latinText).toContain('New');
    expect(latinText).toContain('Roman');
    // The glued non-starter "、" follows "Roman" on that same line (kinsoku keeps
    // it off the next line's head). The line is now "Times New Roman、" — full.
    expect(latinText).toContain('、');
    // No line may BEGIN with the non-starter "、".
    for (const g of lines) expect(g[0].text.startsWith('、')).toBe(false);
  });

  it('flushes "Roman" + "、" together when even "、" has no room (no mid-word split, no orphan)', async () => {
    // Tighten the line so "Times New " uses all but 5 cells: leading CJK fills
    // line 1; line 2 = "Times " (6) + "New " (4) = 10 cells; "Roman" (5) exactly
    // fills to 15, leaving 1 cell — but we shrink the column to 15 cells so that
    // after "Times New " (10) only 5 remain: "Roman" (5) fits exactly with ZERO
    // room for "、". The pre-flush must then fire (Roman+"、" = 6 > 5) so "Roman"
    // and "、" start the next line TOGETHER — never "Roman" split mid-word and
    // never "、" orphaned at a head.
    const calls = await render(
      [gluedPara(['あああああああああああああ', 'Times New Roman', '、あいうえお'])],
      section({ pageWidth: 300 }), // 15 cells/line
    );
    expect(calls.length).toBeGreaterThan(0);

    const lines = linesByY(calls);
    // No line may BEGIN with the non-starter "、".
    for (const g of lines) {
      expect(g[0].text.startsWith('、')).toBe(false);
    }
    // "Roman" is never split across lines: the full word lands on one line, and
    // wherever "Roman" is, "、" is on that same line (glued).
    const romanLine = lines.find((g) => g.map((x) => x.text).join('').includes('Roman'));
    expect(romanLine).toBeDefined();
    const romanText = romanLine!.map((x) => x.text).join('');
    expect(romanText).toContain('、');
    // "Times New" is contiguous Latin and must not be broken inside a word either.
    expect(calls.map((c) => c.text).join('')).toContain('Roman');
  });
});
