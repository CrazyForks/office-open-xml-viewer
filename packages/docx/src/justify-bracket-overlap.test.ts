import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas, type DocxTextRunInfo } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
} from './types';

// ECMA-376 §17.18.44 (ST_Jc `both`/`distribute`) — opening-bracket overlap on the
// JUSTIFY draw path (the sibling of the docGrid case-1 fix, PR #626).
//
// A justified (`jc=both`) paragraph WITHOUT a docGrid charSpace goes through the
// case-2 justify branch of the draw loop. distributeLineSlack opens a gap at
// EVERY inter-CJK boundary on a pure-CJK line, so `stretch.splitBefore` lists a
// cut before each glyph (length === cps.length - 1). The OLD case-2 code then
// drew ONE single-glyph piece per code point via isolated `fillText` calls,
// while positioning glyph i by the CONTEXTUAL cumulative `measureText(prefix_i)`.
//
// The contextual collapse that bites here is JIS X 4051 約物連続 (consecutive-
// punctuation packing), NOT a bracket-next-to-kana collapse: a CLOSING-class glyph
// immediately followed by an OPENING bracket — "：［", "、［", "）（" — packs the
// pair ~half-em tighter in `measureText` (verified on real fonts: "［本" does NOT
// pack, "名：［" does). So the cumulative measure stepping INTO the bracket (after
// "：") is half-width, but the OLD draw painted the bracket ISOLATED at FULL width,
// so it overran its successor by ~half-em — the next glyph was pulled under the
// bracket (the "分 ⊂ ［" smashing seen in sample-16).
//
// The fix: when a run is FULLY distributed (a gap at every internal boundary ⇒
// uniform pitch), draw the whole CONTEXTUALLY-shaped run in ONE `fillText` with
// `ctx.letterSpacing = distPerGap`. measure and draw then shape the SAME
// (contextual) way ⇒ the packing is honoured identically and nothing overlaps;
// the final glyph still lands on the segment box edge (= internalStretch) so the
// next run abuts (verified: drawn extent unchanged, no right-margin overflow).
//
// NOTE: the load-bearing guard below is STRUCTURAL and font-independent — the
// fully-distributed line is drawn as exactly ONE contiguous `fillText` with
// `ctx.letterSpacing` set (vs the OLD N isolated single-glyph draws). The
// `ctxMeasure` mock models the 約物連続 packing only to reconstruct the OLD path's
// overlap illustratively.

const FONT_PX = 20; // glyph advance per full-width CJK char in the stub (scale 1)

// JIS X 4051 約物連続 (consecutive-punctuation packing): a CLOSING-class glyph
// immediately FOLLOWED by an OPENING bracket has its adjacent empty half-bodies
// merged, so the pair measures ~half-em tighter. (A bare opening bracket next to
// a kana/kanji does NOT collapse — verified on real fonts; only the punctuation
// PAIR packs.) These are the classes the browser's measureText compresses.
const CLOSE_PUNCT = '：。）、］」』';
const OPEN_BRACKET = '［（「『';

/** Model the browser's contextual width: full-em per glyph, MINUS a half-em at
 *  each adjacency where a closing-class punctuation is immediately followed by an
 *  opening bracket (約物連続 packing, e.g. "：［", "］。"→no, "、［", "）（"). So
 *  `measure("名：［") < measure("名：") + measure("［")` while `measure("［本")
 *  === measure("［") + measure("本")` (a bracket next to a kanji does NOT pack).
 *  The collapse must NOT include letterSpacing — the renderer sets letterSpacing
 *  AFTER all measureText calls, and the canvas adds it between glyphs. */
function ctxMeasure(s: string, fontPx: number): number {
  const cps = [...s];
  let w = 0;
  for (let i = 0; i < cps.length; i++) {
    w += fontPx;
    if (i > 0 && OPEN_BRACKET.includes(cps[i]) && CLOSE_PUNCT.includes(cps[i - 1])) {
      w -= fontPx / 2; // closing-punct + opening-bracket pair packs half-em tighter
    }
  }
  return w;
}

/** Recording 2D context. measureText models 約物半角 contextual collapse; fillText
 *  records text + x + y AND the current letterSpacing at call time (so we can
 *  assert the fully-distributed fast-path sets `ctx.letterSpacing = distPerGap`). */
function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  fillTextCalls: { text: string; x: number; y: number; letterSpacing: string }[];
} {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const fillTextCalls: { text: string; x: number; y: number; letterSpacing: string }[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(v: string) { letterSpacing = v; },
    measureText: (s: string) => {
      const p = px();
      const w = ctxMeasure(s, p);
      return {
        width: w,
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
    fillText(text: string, x: number, y: number) {
      fillTextCalls.push({ text, x, y, letterSpacing });
    },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0,
    height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, fillTextCalls };
}

function textRun(text: string, fontSize = FONT_PX): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  };
}

type DocRun = DocParagraph['runs'][number];

function para(
  runs: DocxTextRun[],
  opts: {
    alignment?: DocParagraph['alignment'];
    overflowPunct?: boolean;
  } = {},
): BodyElement {
  const p: DocParagraph = {
    alignment: opts.alignment ?? 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: runs.map((r) => ({ type: 'text', ...r }) as DocRun),
    defaultFontSize: FONT_PX, defaultFontFamily: 'NotInMetrics',
    widowControl: false,
    overflowPunct: opts.overflowPunct,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function section(overrides: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 600, pageHeight: 400,
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    // NO docGrid charSpace ⇒ gridSegDeltaPx is 0 ⇒ case 1 is inactive and the
    // justify path (case 2) is taken. This is sample-16's configuration.
    docGridCharSpace: undefined,
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

async function render(
  body: BodyElement[],
  sec: SectionProps,
): Promise<{
  runs: DocxTextRunInfo[];
  fillTextCalls: { text: string; x: number; y: number; letterSpacing: string }[];
}> {
  const { canvas, fillTextCalls } = makeRecordingCanvas();
  const runs: DocxTextRunInfo[] = [];
  await renderDocumentToCanvas(doc(body, sec), canvas, 0, {
    dpr: 1,
    width: sec.pageWidth, // scale = 1 (px per pt)
    onTextRun: (r) => runs.push(r),
  });
  return { runs, fillTextCalls };
}

// A pure-CJK run carrying a real 約物連続 pair: "：［" (closing punctuation
// immediately followed by an opening bracket), long enough to wrap on a ~200px
// page. The first (justified, non-last) line carries the "：［分類］" packing pair.
const TEXT = 'スタイル名：［分類］。あいうえおかきくけこさしすせそたちつてと';
// ~200px page ⇒ ~10 full-width glyphs per line ⇒ the first wrapped line includes
// the "：［" pair and is NOT the last line (so it is justified). Keep the page
// narrow so the line is stretched.
const PAGE_W = 200;

/** Group the recorded fillText calls by their baseline y, sorted top-to-bottom. */
function linesByY(
  calls: { text: string; x: number; y: number; letterSpacing: string }[],
): { y: number; calls: { text: string; x: number; y: number; letterSpacing: string }[] }[] {
  const byY = new Map<number, { text: string; x: number; y: number; letterSpacing: string }[]>();
  for (const c of calls) {
    const key = Math.round(c.y);
    let arr = byY.get(key);
    if (!arr) { arr = []; byY.set(key, arr); }
    arr.push(c);
  }
  return [...byY.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([y, cs]) => ({ y, calls: cs.slice().sort((p, q) => p.x - q.x) }));
}

describe('justified paragraph (no docGrid) — 約物連続 brackets are drawn contiguously, never overlap (§17.18.44)', () => {
  // (1) RED→GREEN: the first (justified, non-last) line's pure-CJK segment is
  //     painted as ONE contiguous fillText carrying the whole line text, with
  //     ctx.letterSpacing === `${distPerGap}px` (a non-zero px). Before the fix:
  //     many single-codepoint draws with letterSpacing '0px'.
  it('draws the fully-distributed justify line as a single contiguous fillText with letterSpacing=distPerGap', async () => {
    const { fillTextCalls } = await render(
      [para([textRun(TEXT)], { alignment: 'both', overflowPunct: false })],
      section({ pageWidth: PAGE_W }),
    );
    expect(fillTextCalls.length).toBeGreaterThan(0);

    const lines = linesByY(fillTextCalls);
    expect(lines.length).toBeGreaterThanOrEqual(2); // wrapped
    const first = lines[0];

    // GREEN: the justified first line is ONE fillText (not per-code-point).
    expect(first.calls.length, 'one contiguous fillText for the first justified line').toBe(1);
    const drawn = first.calls[0];
    // It carries the bracket and is NOT a single code point.
    expect([...drawn.text].length, 'the contiguous draw spans many code points').toBeGreaterThan(1);
    expect(drawn.text).toContain('［');

    // letterSpacing carries the justify pitch (non-zero px), not '0px'.
    expect(drawn.letterSpacing).toMatch(/^-?\d+(\.\d+)?px$/);
    expect(drawn.letterSpacing).not.toBe('0px');
    const distPerGap = parseFloat(drawn.letterSpacing);
    expect(distPerGap, 'justify pitch is a positive expansion').toBeGreaterThan(0);
  });

  // (2) No-overlap invariant: with the contiguous draw there is exactly ONE call
  //     for the justified line, at the line's left start; the browser places the
  //     intra-line glyph positions, so no "分 ⊂ ［" overlap can occur. We also
  //     reconstruct the per-glyph positions the OLD isolated path would have used
  //     and assert the bracket-overlap it produced is GONE.
  it('keeps the opening bracket from overlapping the next glyph (the 分⊂［ regression is gone)', async () => {
    const { runs, fillTextCalls } = await render(
      [para([textRun(TEXT)], { alignment: 'both', overflowPunct: false })],
      section({ pageWidth: PAGE_W }),
    );

    const lines = linesByY(fillTextCalls);
    const first = lines[0];
    // ONE contiguous draw for the line ⇒ no isolated per-glyph placement.
    expect(first.calls.length).toBe(1);
    const drawn = first.calls[0];

    // The single draw starts at the line's left start (the reported run x), and
    // the browser handles intra-line glyph positions, so by construction no glyph
    // is painted left of the previous glyph's end.
    const lineRun = runs.find((r) => r.text === drawn.text || drawn.text.startsWith(r.text));
    if (lineRun) expect(drawn.x).toBeCloseTo(lineRun.x, 3);

    // Cross-check the regression (illustrative model): reconstruct where the OLD
    // isolated per-code-point path drew each glyph (dx = CONTEXTUAL
    // measure(prefix_i) + i·distPerGap) and confirm that path produced a real
    // overlap — and that the fix (a single contiguous draw) makes it impossible.
    const lineText = drawn.text;
    const cps = [...lineText];
    const distPerGap = parseFloat(drawn.letterSpacing);
    expect(lineText).toContain('［');
    // OLD-path x positions (the overlap source): isolated glyph i at
    // measure(prefix_i) + i·distPerGap. The 約物連続 packing collapses the "：［"
    // pair, so the cumulative measure stepping past the bracket is half-em short,
    // pulling the glyph AFTER the bracket left of where the FULL-width isolated
    // bracket actually ends.
    const oldXs = cps.map((_c, i) =>
      ctxMeasure(cps.slice(0, i).join(''), FONT_PX) + i * distPerGap,
    );
    // Each glyph was painted ISOLATED, so it occupies a FULL FONT_PX advance from
    // its x. Scan adjacent pairs for a glyph whose isolated full-width end exceeds
    // the next glyph's start ⇒ OVERLAP (this is the "分⊂［" the real-font probe
    // measured at ~half-em / 7.4px on sample-16's "：［" pairs).
    let worstOverlap = 0;
    for (let i = 1; i < cps.length; i++) {
      const prevIsolatedEnd = oldXs[i - 1] + FONT_PX;
      const overlap = prevIsolatedEnd - oldXs[i];
      if (overlap > worstOverlap) worstOverlap = overlap;
    }
    expect(
      worstOverlap,
      'OLD isolated per-code-point path overlapped at the 約物連続 "：［" pair',
    ).toBeGreaterThan(1); // ~half-em (≈ FONT_PX/2) minus one gap pitch

    // The fix: we drew ONE contiguous fillText, so the browser shapes the whole
    // run the SAME way it was measured (約物半角 collapse honoured) ⇒ no isolated
    // glyph can be pulled left of its neighbour. There is exactly ONE draw for the
    // line, carrying the full wrapped line text at its left start — the browser
    // owns the intra-line glyph positions, so the overlap is structurally GONE.
    expect(first.calls.length).toBe(1);
    expect(drawn.text).toBe(lineText);
  });

  // (3) Last line / non-justified unaffected.
  it('leaves the paragraph last line and a left-aligned control un-justified (single draw, 0px)', async () => {
    // (a) The justified paragraph's LAST line is not justified ⇒ single fillText,
    //     letterSpacing '0px'.
    const both = await render(
      [para([textRun(TEXT)], { alignment: 'both' })],
      section({ pageWidth: PAGE_W }),
    );
    const bothLines = linesByY(both.fillTextCalls);
    const last = bothLines[bothLines.length - 1];
    expect(last.calls.length, 'last line is a single un-justified draw').toBe(1);
    expect(last.calls[0].letterSpacing, 'last line not stretched').toBe('0px');

    // (b) A left-aligned control paragraph is unchanged: every line a single
    //     contiguous draw at letterSpacing '0px'.
    const left = await render(
      [para([textRun(TEXT)], { alignment: 'left' })],
      section({ pageWidth: PAGE_W }),
    );
    const leftLines = linesByY(left.fillTextCalls);
    for (const ln of leftLines) {
      expect(ln.calls.length, 'left-aligned line drawn as one contiguous fillText').toBe(1);
      expect(ln.calls[0].letterSpacing, 'left-aligned line not stretched').toBe('0px');
    }
  });

  // (4) Partial-gap path unaffected: a justified line where splitBefore is a
  //     SUBSET of the inter-glyph boundaries (length < cps.length - 1) must keep
  //     the piece loop (multiple fillText). A mixed CJK+Latin run produces gaps
  //     only at inter-CJK boundaries (not inside the Latin word), so the line's
  //     CJK segment is NOT fully distributed and draws as multiple pieces.
  it('keeps the multi-glyph piece loop on a partial-gap (CJK+Latin) justified line', async () => {
    // CJK …, then a Latin word (no internal gaps), then CJK — long enough to wrap.
    // The justified first line mixes a CJK run with a contiguous Latin token, so
    // not every inter-glyph boundary opens a gap.
    const MIXED = 'あいうえおAbcdefgかきくけこさしすせそたちつてと';
    const { fillTextCalls } = await render(
      [para([textRun(MIXED)], { alignment: 'both' })],
      section({ pageWidth: PAGE_W }),
    );
    const lines = linesByY(fillTextCalls);
    expect(lines.length).toBeGreaterThanOrEqual(2);
    const first = lines[0];
    // The partial-gap line draws MORE than one fillText (the piece loop), because
    // the Latin token is a contiguous piece while inter-CJK boundaries are cut.
    // (If the wrap happened to leave a fully-CJK first line, this still documents
    //  the fully-CJK fast-path; assert at least that a justified line with a Latin
    //  token uses multiple draws somewhere.)
    const anyMultiPieceJustified = lines.some(
      (ln) => ln.calls.length > 1 && ln.calls.some((c) => /[A-Za-z]/.test(c.text)),
    );
    expect(
      anyMultiPieceJustified,
      'a justified line containing a Latin token uses the multi-piece loop',
    ).toBe(true);
  });
});
