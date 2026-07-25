import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types';

// Interop behaviour (NOT ECMA-376 — see isSectionBreakSpacerAt): an EMPTY
// paragraph that carries a section break (an inkless paragraph immediately
// followed by a `sectionBreak` element) has its spacing-BEFORE suppressed. Word
// and LibreOffice both render such a "section-break spacer" flush below the
// preceding paragraph (sample-13: the empty `mSectionBreak` between "Keywords"
// and "1. INTRODUCTION" carries before=22pt but neither editor applies it). A
// normal empty paragraph (not followed by a section break) keeps its before.

interface Call { text: string; y: number; }

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; calls: Call[] } {
  let font = '10px serif';
  const calls: Call[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(s: string, _x: number, y: number) { calls.push({ text: s, y }); },
    strokeText(s: string, _x: number, y: number) { calls.push({ text: s, y }); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls };
}

function para(text: string, spaceBefore = 0, spaceAfter = 0): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore, spaceAfter, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: text
      ? [{
          type: 'text', text, bold: false, italic: false, underline: false,
          strikethrough: false, fontSize: 10, color: null, fontFamily: 'Times New Roman',
          fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
        } as DocParagraph['runs'][number]]
      : [],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

const MATCHING_SECTION_GEOMETRY = {
  pageWidth: 400, pageHeight: 600,
  marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
  headerDistance: 4, footerDistance: 4,
};

function docOf(body: BodyElement[]): DocxDocumentModel {
  const section: SectionProps = {
    ...MATCHING_SECTION_GEOMETRY, titlePage: false, evenAndOddHeaders: false,
    // The final section (containing B) starts CONTINUOUS so the section break is
    // not a page break and B stays on page 0 (§17.6.22 — the break is governed by
    // the FOLLOWING section's start type).
    sectionStart: 'continuous',
  } as SectionProps;
  return {
    section,
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

async function baselineOf(body: BodyElement[], text: string): Promise<number> {
  const { canvas, calls } = makeRecordingCanvas();
  await renderDocumentToCanvas(docOf(body), canvas, 0, { dpr: 1, width: 400 });
  const c = calls.find((k) => k.text === text);
  expect(c, `expected to paint ${text}`).toBeDefined();
  return (c as Call).y;
}

const SPACER_BEFORE = 20;

describe('section-break spacer suppresses spacing-before (Word/LibreOffice interop)', () => {
  it('an empty paragraph followed by a section break drops its 20pt before', async () => {
    // [A] [empty before=20] [sectionBreak continuous] [B]
    const spacerBody: BodyElement[] = [
      para('A') as unknown as BodyElement,
      para('', SPACER_BEFORE) as unknown as BodyElement,
      { type: 'sectionBreak', kind: 'continuous', geom: MATCHING_SECTION_GEOMETRY } as unknown as BodyElement,
      para('B') as unknown as BodyElement,
    ];
    // Control: same, but the empty paragraph is NOT followed by a section break,
    // so it is a normal empty paragraph and keeps its 20pt before.
    const controlBody: BodyElement[] = [
      para('A') as unknown as BodyElement,
      para('', SPACER_BEFORE) as unknown as BodyElement,
      para('B') as unknown as BodyElement,
    ];

    const aSpacer = await baselineOf(spacerBody, 'A');
    const bSpacer = await baselineOf(spacerBody, 'B');
    const aControl = await baselineOf(controlBody, 'A');
    const bControl = await baselineOf(controlBody, 'B');

    // 'A' is at the same place in both (nothing above it changed).
    expect(aSpacer).toBeCloseTo(aControl, 3);
    // The control applies the empty paragraph's 20pt before; the section-break
    // spacer suppresses it, so B sits exactly 20pt higher.
    expect(bControl - bSpacer).toBeCloseTo(SPACER_BEFORE, 1);
  });

  it('suppresses when the break MARKER is nextPage but the section resolves continuous (§17.6.22 — sample-13 shape)', async () => {
    // sample-13: the sectPr-bearing spacer's break marker is kind="nextPage", yet
    // the FOLLOWING section starts continuous (docOf sets section.sectionStart),
    // so the EFFECTIVE break is continuous and the spacer's before is dropped. Two
    // empty paragraphs like the real doc; the SECOND (immediately before the break)
    // is the spacer whose before collapses against the first.
    const body: BodyElement[] = [
      para('A') as unknown as BodyElement,
      para('', SPACER_BEFORE) as unknown as BodyElement, // empty 1
      para('', SPACER_BEFORE) as unknown as BodyElement, // empty 2 = spacer
      // Author the outgoing page box explicitly so this test isolates the
      // resolved continuous boundary; §17.6.22 does not let a non-continuous
      // section inherit omitted geometry from the following section.
      { type: 'sectionBreak', kind: 'nextPage', geom: MATCHING_SECTION_GEOMETRY } as unknown as BodyElement,
      para('B') as unknown as BodyElement,
    ];
    const control: BodyElement[] = [
      para('A') as unknown as BodyElement,
      para('', SPACER_BEFORE) as unknown as BodyElement,
      para('', SPACER_BEFORE) as unknown as BodyElement,
      para('B') as unknown as BodyElement,
    ];
    const bBody = await baselineOf(body, 'B');
    const bControl = await baselineOf(control, 'B');
    // …but the section resolves continuous, so the SECOND empty's before is dropped
    // (gap empty1→empty2 = 0): B is exactly 20pt higher than the no-break control.
    expect(bControl - bBody).toBeCloseTo(SPACER_BEFORE, 1);
  });

  it('a NON-empty paragraph followed by a section break keeps its before (only empty spacers are suppressed)', async () => {
    // The section-ending paragraph here has text, so it is not an inkless spacer.
    const withText: BodyElement[] = [
      para('A') as unknown as BodyElement,
      para('X', SPACER_BEFORE) as unknown as BodyElement,
      { type: 'sectionBreak', kind: 'continuous', geom: MATCHING_SECTION_GEOMETRY } as unknown as BodyElement,
      para('B') as unknown as BodyElement,
    ];
    const control: BodyElement[] = [
      para('A') as unknown as BodyElement,
      para('X', SPACER_BEFORE) as unknown as BodyElement,
      para('B') as unknown as BodyElement,
    ];
    const xWith = await baselineOf(withText, 'X');
    const xControl = await baselineOf(control, 'X');
    // The non-empty paragraph keeps its before regardless of the following break.
    expect(xWith).toBeCloseTo(xControl, 1);
  });
});

describe('collapsed continuous-section spacer — Word section-mark collapse (sample-12)', () => {
  // NEW-a (isCollapsedContinuousSpacer): a continuous-section spacer with NO
  // space-before of its own renders NO paragraph-mark line box (sample-12 — Word
  // shows ONE blank line, not two, and the heading sits ~24pt higher). A spacer
  // WITH a space-before keeps its box (sample-13). Reconstructed from Word's output.
  it('a zero-before continuous spacer drops its mark line; a non-zero-before one keeps it', async () => {
    const collapse: BodyElement[] = [
      para('A') as unknown as BodyElement,
      para('') as unknown as BodyElement, // spacer before=0 → collapses (no line box)
      { type: 'sectionBreak', kind: 'continuous', geom: MATCHING_SECTION_GEOMETRY } as unknown as BodyElement,
      para('B') as unknown as BodyElement,
    ];
    const keepBox: BodyElement[] = [
      para('A') as unknown as BodyElement,
      para('', SPACER_BEFORE) as unknown as BodyElement, // spacer before>0 → keeps its line box
      { type: 'sectionBreak', kind: 'continuous', geom: MATCHING_SECTION_GEOMETRY } as unknown as BodyElement,
      para('B') as unknown as BodyElement,
    ];
    const aCollapse = await baselineOf(collapse, 'A');
    const bCollapse = await baselineOf(collapse, 'B');
    const aKeep = await baselineOf(keepBox, 'A');
    const bKeep = await baselineOf(keepBox, 'B');
    // 'A' is unchanged in both (nothing above it differs).
    expect(aCollapse).toBeCloseTo(aKeep, 3);
    // The kept-box spacer adds one mark line; the collapsed one adds none — so B is
    // strictly higher when the spacer collapses (by ~one line height). The spacer's
    // own before is suppressed in BOTH cases, so the difference is purely the box.
    expect(bKeep - bCollapse).toBeGreaterThan(6);
  });

  // NEW-b (leadsCollapsedRun): the empty paragraph that begins the section-break
  // run (immediately before the collapsed spacer) sits FLUSH below the preceding
  // paragraph — its space-after is dropped (sample-12: "[Format…]"'s 6pt after
  // vanishes, placing the heading at Word's 446pt). No-op when not collapsing.
  it('drops the previous paragraph space-after when an empty run leads into a collapsed spacer', async () => {
    const mk = (aAfter: number): BodyElement[] => [
      para('A', 0, aAfter) as unknown as BodyElement, // A carries a space-after
      para('') as unknown as BodyElement, // empty-run start (inkless), leads the spacer
      para('') as unknown as BodyElement, // spacer before=0 → collapses
      { type: 'sectionBreak', kind: 'continuous', geom: MATCHING_SECTION_GEOMETRY } as unknown as BodyElement,
      para('B') as unknown as BodyElement,
    ];
    const bNoAfter = await baselineOf(mk(0), 'B');
    const bBigAfter = await baselineOf(mk(40), 'B');
    // A's 40pt space-after is dropped at the section-break run, so B does NOT move
    // down — it stays within rounding of the no-after case.
    expect(bBigAfter - bNoAfter).toBeCloseTo(0, 1);
  });
});
