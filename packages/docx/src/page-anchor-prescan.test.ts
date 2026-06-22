import { describe, expect, it } from 'vitest';
import {
  __test_isPageLevelAnchorY,
  __test_preRegisterPageFloats,
  computePages,
  type RenderState,
} from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  ImageRun,
  SectionProps,
  ShapeRun,
} from './types.js';

// Pins ECMA-376 §20.4.3.2 / §20.4.3.5 behaviour: a `<wp:anchor>` whose
// `<wp:positionV>@relativeFrom` resolves Y independently of the source-order
// anchoring paragraph (page / margin / *Margin / column) is laid out by Word
// as soon as the page is opened, so paragraphs that PRECEDE the anchor's
// paragraph in source order still wrap around the float. Renderer mirrors
// this by pre-scanning page-level floats at every page-start
// (renderBodyElements + paginator's float-reset call sites) and recording
// the source paragraph in `state.pageAnchorPrescanned` so the per-paragraph
// `registerAnchorFloats` skips the duplicate registration when the flow
// reaches it. These three tests cover (1) the wrap-window narrowing seen by
// an EARLIER paragraph, (2) the carve-out for paragraph-local Y (which must
// NOT pre-register), and (3) idempotency (no double FloatRect).

// ===== Stub canvas + helpers (mirrors pagination.test) =====

function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
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
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, textAlign: 'left' as CanvasTextAlign,
    direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

function section(overrides: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 200, pageHeight: 200,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    ...overrides,
  };
}

type DocRun = DocParagraph['runs'][number];

function textRun(text: string, fontSize: number): DocRun {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  };
  return { type: 'text', ...run } as DocRun;
}

function para(opts: { text?: string; fontSize?: number } = {}): BodyElement {
  const fontSize = opts.fontSize ?? 20;
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: opts.text ? [textRun(opts.text, fontSize)] : [],
    defaultFontSize: fontSize, defaultFontFamily: 'NotInMetrics',
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function paraWith(runs: DocRun[], opts: { fontSize?: number } = {}): BodyElement {
  const fontSize = opts.fontSize ?? 20;
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs,
    defaultFontSize: fontSize, defaultFontFamily: 'NotInMetrics',
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

// Minimal page-level square-anchored image float (positionV relativeFrom=margin
// + align=top, square wrap). 100×40 pt, pinned at the right edge of the top
// margin so anything on the page from y∈[20,60], x∈[120,220] wraps around it.
function pageImageRun(opts: {
  widthPt?: number;
  heightPt?: number;
  anchorXPt?: number;
  anchorYPt?: number;
  anchorXRelativeFrom?: string | null;
  anchorYRelativeFrom?: string | null;
  anchorXFromMargin?: boolean;
  anchorYFromPara?: boolean;
  wrapMode?: string;
  wrapSide?: string;
} = {}): DocRun {
  const img: ImageRun = {
    imagePath: 'word/media/test1.png',
    mimeType: 'image/png',
    widthPt: opts.widthPt ?? 60,
    heightPt: opts.heightPt ?? 40,
    anchor: true,
    anchorXPt: opts.anchorXPt ?? 100,
    anchorYPt: opts.anchorYPt ?? 0,
    anchorXFromMargin: opts.anchorXFromMargin ?? true,
    anchorYFromPara: opts.anchorYFromPara ?? false,
    wrapMode: opts.wrapMode ?? 'square',
    wrapSide: opts.wrapSide ?? 'bothSides',
    anchorXRelativeFrom: opts.anchorXRelativeFrom ?? 'margin',
    anchorYRelativeFrom: opts.anchorYRelativeFrom ?? 'margin',
  };
  return { type: 'image', ...img } as DocRun;
}

// Minimal page-level wrap shape (positionV relativeFrom=page + topAndBottom).
function pageShapeRun(opts: {
  widthPt?: number;
  heightPt?: number;
  anchorYPt?: number;
  anchorYRelativeFrom?: string | null;
  anchorYFromPara?: boolean;
  wrapMode?: string | null;
} = {}): DocRun {
  const s: ShapeRun = {
    widthPt: opts.widthPt ?? 160,
    heightPt: opts.heightPt ?? 30,
    anchorXPt: 0,
    anchorYPt: opts.anchorYPt ?? 20,
    anchorXFromMargin: true,
    anchorYFromPara: opts.anchorYFromPara ?? false,
    zOrder: 0,
    subpaths: [],
    presetGeometry: 'rect',
    fill: { fillType: 'solid', color: 'FFFFFF' },
    stroke: null,
    wrapMode: opts.wrapMode === undefined ? 'topAndBottom' : opts.wrapMode,
    wrapSide: null,
    distTop: 0, distBottom: 0, distLeft: 0, distRight: 0,
    anchorYRelativeFrom: opts.anchorYRelativeFrom ?? 'page',
  } as ShapeRun;
  return { type: 'shape', ...s } as DocRun;
}

// ===== Tests =====

describe('preRegisterPageFloats — isPageLevelAnchorY classifier (§20.4.3.5)', () => {
  it('paragraph/line/character ⇒ paragraph-local (NOT page-level)', () => {
    expect(__test_isPageLevelAnchorY('paragraph', false)).toBe(false);
    expect(__test_isPageLevelAnchorY('line', false)).toBe(false);
    expect(__test_isPageLevelAnchorY('character', false)).toBe(false);
  });

  it('page / margin / *Margin / column ⇒ page-level', () => {
    for (const rf of ['page', 'margin', 'topMargin', 'bottomMargin', 'leftMargin', 'rightMargin', 'insideMargin', 'outsideMargin', 'column']) {
      expect(__test_isPageLevelAnchorY(rf, false)).toBe(true);
    }
  });

  it('absent relativeFrom defers to anchorYFromPara (page-level only when NOT from-para)', () => {
    expect(__test_isPageLevelAnchorY(null, false)).toBe(true);
    expect(__test_isPageLevelAnchorY(null, true)).toBe(false);
    expect(__test_isPageLevelAnchorY(undefined, false)).toBe(true);
  });
});

describe('preRegisterPageFloats — earlier paragraph sees the float band (§20.4.3.2/§20.4.3.5)', () => {
  // Page-content geometry: contentH = 200 - 40 = 160, contentW = 200 - 40 = 160.
  // A page-anchored topAndBottom shape carried by paragraph B is pinned at
  // y∈[20,50] (anchorYPt=20 + heightPt=30, anchorYRelativeFrom='page' ⇒ Y is
  // page-absolute; a full-width topAndBottom band stamps the full content
  // width). Paragraph A's first body line WITHOUT pre-scan starts at the
  // marginTop (no float band registered yet), so its text fits in column-top
  // space. WITH pre-scan, paragraph A is laid out AFTER the float band is
  // already in the float set ⇒ its lines flow BELOW y=50 (under the band),
  // displacing the rest of the body downward.
  //
  // Asserts the displacement by comparing the paginated layout against a
  // sibling test with a paragraph-LOCAL anchor (relativeFrom='paragraph'),
  // which is NOT pre-registered, so paragraph A keeps the full content top.

  it('a page-level anchor on a LATER paragraph wraps earlier paragraphs (more pages)', () => {
    // Geometry: contentW=160, contentH=160. 20pt font → 8 chars/line normal.
    // Paragraph B carries a page-level SQUARE float anchored at x=80 (right
    // half of content band) y=0 in the top margin, width 80, height 60. Its
    // exclusion rect (without dist padding) is x∈[100,180], y∈[20,80] (the
    // anchorXFromMargin/anchorYRelativeFrom='margin' resolve to marginLeft+80
    // = 100, marginTop+0 = 20). For lines whose top is in [20,80]
    // (3 lines: y=20→40, 40→60, 60→80), the wrap window narrows to
    // [20,100] = 80pt = 4 chars/line. Below the float, lines use the full
    // 160pt = 8 chars/line.
    //
    // 64 chars of text in paragraph A — exactly 8 lines × 20pt = 160pt
    //   pre-scan OFF (paragraph-local Y): A flows full-width ⇒ 8 lines fit
    //                exactly on page 1; B's anchor-only mark also fits
    //                (spaceAfter overflow allowed) ⇒ ONE page.
    //   pre-scan ON: lines 1-3 (y∈[20,80]) wrap into x∈[20,100] = 80pt =
    //                4 chars/line ⇒ 12 chars consumed in those 3 lines.
    //                Remaining 52 chars below the float at 8 chars/line ⇒
    //                ceil(52/8) = 7 lines. Total 10 lines × 20 = 200pt > 160pt
    //                ⇒ paragraph A alone spills onto a second page.
    const bodyPage = [
      para({ text: 'あ'.repeat(64), fontSize: 20 }),
      paraWith([
        pageImageRun({
          widthPt: 80, heightPt: 60,
          anchorXPt: 80, anchorYPt: 0,
          anchorXRelativeFrom: 'margin', anchorYRelativeFrom: 'margin',
          anchorXFromMargin: true, anchorYFromPara: false,
          wrapMode: 'square', wrapSide: 'bothSides',
        }),
      ]),
    ];
    const pagesWithPageAnchor = computePages(bodyPage, section(), makeCtx());

    // Paragraph-LOCAL Y anchor (relativeFrom='paragraph'). Pre-scan must
    // SKIP this float — A flows without it.
    const bodyLocal = [
      para({ text: 'あ'.repeat(64), fontSize: 20 }),
      paraWith([
        pageImageRun({
          widthPt: 80, heightPt: 60,
          anchorXPt: 80, anchorYPt: 0,
          anchorXRelativeFrom: 'margin', anchorYRelativeFrom: 'paragraph',
          anchorXFromMargin: true, anchorYFromPara: true,
          wrapMode: 'square', wrapSide: 'bothSides',
        }),
      ]),
    ];
    const pagesWithLocalAnchor = computePages(bodyLocal, section(), makeCtx());

    // Signal: under the pre-scan, paragraph A is SPLIT (its 64 chars no
    // longer fit on one page because the pre-registered float narrows the
    // wrap window) — page 1 carries a `lineSlice` for A. Under the
    // paragraph-local anchor, A is NOT split (8 lines × 20pt = 160pt fits
    // page 1 exactly) — its element is emitted whole, no `lineSlice`.
    const firstElOf = (pages: ReturnType<typeof computePages>) => pages[0][0] as { lineSlice?: { start: number; end: number } };
    const pageFirst = firstElOf(pagesWithPageAnchor);
    const localFirst = firstElOf(pagesWithLocalAnchor);
    // Page-level case: paragraph A is split (lineSlice present) because the
    // pre-registered float forced more lines than the page can hold.
    expect(pageFirst.lineSlice).toBeDefined();
    // Paragraph-local case: A is NOT pre-narrowed, fits whole on page 1.
    expect(localFirst.lineSlice).toBeUndefined();
  });
});

describe('preRegisterPageFloats — paragraph-local Y is NOT pre-registered', () => {
  it('pre-scan ignores a wrap float with positionV relativeFrom="paragraph"', () => {
    const body = [
      para({ text: 'ABC' }),
      paraWith([
        // Paragraph-local: must NOT be pre-registered (paraId stays at the
        // paragraph's own registration in registerAnchorFloats).
        pageShapeRun({ widthPt: 100, heightPt: 30, anchorYRelativeFrom: 'paragraph', anchorYFromPara: true, wrapMode: 'topAndBottom' }),
      ]),
    ];
    // Minimal RenderState stub matching what registerImageFloat/registerShapeFloat
    // touch: scale, marginLeft/Top, pageH, contentX/W, dryRun (no images), floats.
    const state = {
      scale: 1,
      marginLeft: 20, marginRight: 20, marginTop: 20, marginBottom: 20,
      pageWidth: 200, pageH: 200,
      contentX: 20, contentW: 160,
      floats: [],
      floatParaSeq: 0,
      dryRun: true,
      pageAnchorPrescanned: new Set(),
    } as unknown as RenderState;
    __test_preRegisterPageFloats(body, 0, state);
    expect(state.floats.length).toBe(0);
    expect(state.pageAnchorPrescanned?.size ?? 0).toBe(0);
  });

  it('pre-scan REGISTERS a page-level (relativeFrom="margin") wrap float on an earlier-scanned paragraph', () => {
    const body = [
      para({ text: 'ABC' }), // paragraph A, no floats
      paraWith([
        // Paragraph B carries a page-level square anchor
        pageImageRun({ widthPt: 60, heightPt: 40, anchorYRelativeFrom: 'margin', anchorYFromPara: false, wrapMode: 'square' }),
      ]),
    ];
    const state = {
      scale: 1,
      marginLeft: 20, marginRight: 20, marginTop: 20, marginBottom: 20,
      pageWidth: 200, pageH: 200,
      contentX: 20, contentW: 160,
      floats: [],
      floatParaSeq: 0,
      images: new Map(),
      dryRun: true,
      pageAnchorPrescanned: new Set(),
    } as unknown as RenderState;
    __test_preRegisterPageFloats(body, 0, state);
    expect(state.floats.length).toBe(1);
    expect(state.pageAnchorPrescanned?.size).toBe(1);
  });
});

describe('preRegisterPageFloats — idempotent on repeated pre-scan', () => {
  it('a paragraph already in pageAnchorPrescanned is not re-registered', () => {
    const paraB = paraWith([
      pageImageRun({ widthPt: 60, heightPt: 40, anchorYRelativeFrom: 'margin', wrapMode: 'square' }),
    ]);
    const body = [para({ text: 'A' }), paraB];
    const state = {
      scale: 1,
      marginLeft: 20, marginRight: 20, marginTop: 20, marginBottom: 20,
      pageWidth: 200, pageH: 200,
      contentX: 20, contentW: 160,
      floats: [],
      floatParaSeq: 0,
      images: new Map(),
      dryRun: true,
      pageAnchorPrescanned: new Set(),
    } as unknown as RenderState;
    __test_preRegisterPageFloats(body, 0, state);
    const after1 = state.floats.length;
    expect(after1).toBe(1);
    // Second call: same body, same start index. The dedupe key is the
    // already-pre-registered paragraph in pageAnchorPrescanned, so floats
    // must NOT grow.
    __test_preRegisterPageFloats(body, 0, state);
    expect(state.floats.length).toBe(after1);
  });

  it('stops at the next pageBreak — items after the break are NOT pre-scanned for this page', () => {
    const body = [
      paraWith([pageImageRun({ widthPt: 60, heightPt: 40, anchorYRelativeFrom: 'margin', wrapMode: 'square' })]), // would register
      { type: 'pageBreak' } as BodyElement,                                                                       // page boundary
      paraWith([pageImageRun({ widthPt: 60, heightPt: 40, anchorYRelativeFrom: 'margin', wrapMode: 'square' })]), // must be SKIPPED — belongs to the next page
    ];
    const state = {
      scale: 1,
      marginLeft: 20, marginRight: 20, marginTop: 20, marginBottom: 20,
      pageWidth: 200, pageH: 200,
      contentX: 20, contentW: 160,
      floats: [],
      floatParaSeq: 0,
      images: new Map(),
      dryRun: true,
      pageAnchorPrescanned: new Set(),
    } as unknown as RenderState;
    __test_preRegisterPageFloats(body, 0, state);
    expect(state.floats.length).toBe(1);
    expect(state.pageAnchorPrescanned?.size).toBe(1);
  });
});
