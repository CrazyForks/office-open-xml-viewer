import { describe, expect, it } from 'vitest';
import { createLayoutServices, layoutDocument } from '../renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from '../types.js';

function measureContext(): CanvasRenderingContext2D {
  return {
    font: '10px serif', letterSpacing: '0px', fontKerning: 'auto',
    measureText: (text: string) => ({
      width: [...text].length * 5,
      actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}

function paragraph(): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{
      type: 'text', text: '1', bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'serif',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: 'super',
      hyperlink: null, noteRef: { kind: 'footnote', id: '7' },
    }],
    defaultFontSize: 10, defaultFontFamily: 'serif', widowControl: false,
  } as unknown as DocParagraph;
}

function ordinaryParagraph(text: string): DocParagraph {
  const result = paragraph();
  const run = result.runs[0];
  if (run?.type === 'text') {
    run.text = text;
    run.noteRef = undefined;
    run.vertAlign = null;
  }
  return result;
}

function ordinaryBodyParagraph(text: string): BodyElement {
  return ordinaryParagraph(text) as unknown as BodyElement;
}

function sectionBreak(): BodyElement {
  return {
    type: 'sectionBreak', kind: 'continuous', columns: null,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    titlePage: false,
  } as unknown as BodyElement;
}

function sectionBoundaryModel(
  before: BodyElement,
  startType: 'continuous' | 'nextPage',
): DocxDocumentModel {
  const section = {
    pageWidth: 200, pageHeight: 100,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 5, footerDistance: 5, titlePage: false,
    evenAndOddHeaders: false, sectionStart: startType, columns: null,
  } as SectionProps;
  return {
    section,
    body: [before, sectionBreak(), ordinaryBodyParagraph('after')],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    footnotes: [], endnotes: [], fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

function floatingTable(): BodyElement {
  return {
    type: 'table', colWidths: [40],
    rows: [{
      cells: [{
        content: [paragraph()], colSpan: 1, vMerge: null,
        borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
        background: null, vAlign: 'top', widthPt: 40,
      }],
      rowHeight: null, rowHeightRule: 'auto', isHeader: false,
    }],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left', layout: 'fixed', overlap: 'overlap',
    tblpPr: {
      leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
      horzAnchor: 'text', horzSpecified: true, vertAnchor: 'text',
      tblpX: 20, tblpY: 10,
    },
  } as unknown as BodyElement;
}

function ordinaryTable(
  justification: 'left' | 'center' | 'right',
  indentPt: number,
  bidiVisual: boolean,
): BodyElement {
  const source = floatingTable() as Extract<BodyElement, { type: 'table' }>;
  return {
    ...source,
    jc: justification,
    tblInd: indentPt,
    bidiVisual,
    tblpPr: null,
  } as unknown as BodyElement;
}

function fragmentLineAdvancesPt(fragment: Extract<ReturnType<typeof layoutDocument>['pages'][number]['layers']['body'][number], { kind: 'paragraph' }>): number {
  return fragment.lines.reduce((sum, line) => sum + line.advancePt, 0);
}

describe('canonical producer with a real document model', () => {
  it('composes compatible continuous sections as disjoint regions on one physical page', () => {
    const model = sectionBoundaryModel(ordinaryBodyParagraph('before'), 'continuous');
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const page = layout.pages[0]!;
    const [outgoing, incoming] = page.sectionRegions;

    expect(layout.pages).toHaveLength(1);
    expect(page.sectionRegions).toHaveLength(2);
    expect(outgoing!.blockEndPt).toBe(incoming!.blockStartPt);
    expect(outgoing!.blockEndPt).toBeGreaterThan(outgoing!.blockStartPt);
    expect(page.layers.body.map((node) => node.flowDomainId)).toEqual([
      outgoing!.flowDomainIds[0], incoming!.flowDomainIds[0],
    ]);
  });

  it('retains a floating table in an empty outgoing region without charging incoming flow', () => {
    const model = sectionBoundaryModel(floatingTable(), 'continuous');
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const page = layout.pages[0]!;
    const [outgoing, incoming] = page.sectionRegions;
    const [floating, follower] = page.layers.body;

    expect(layout.pages).toHaveLength(1);
    expect(page.sectionRegions).toHaveLength(2);
    expect(outgoing!.blockStartPt).toBe(outgoing!.blockEndPt);
    expect(outgoing!.blockEndPt).toBe(incoming!.blockStartPt);
    expect(floating).toMatchObject({
      kind: 'table', ordinaryFlow: false, flowDomainId: outgoing!.flowDomainIds[0],
    });
    expect(follower).toMatchObject({
      kind: 'paragraph', ordinaryFlow: true, flowDomainId: incoming!.flowDomainIds[0],
      flowBounds: { yPt: incoming!.blockStartPt },
    });
    expect(page.flowDomains.find((domain) => domain.id === outgoing!.flowDomainIds[0]))
      .toMatchObject({ logicalBounds: { heightPt: 0 } });
  });

  it('keeps a non-continuous control on separate physical pages', () => {
    const model = sectionBoundaryModel(floatingTable(), 'nextPage');
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.pages).toHaveLength(2);
    expect(layout.pages.map((page) => page.sectionRegions.length)).toEqual([1, 1]);
    expect(layout.pages[0]!.layers.body[0]).toMatchObject({
      kind: 'table', ordinaryFlow: false,
      flowDomainId: layout.pages[0]!.sectionRegions[0]!.flowDomainIds[0],
    });
    expect(layout.pages[1]!.layers.body[0]).toMatchObject({
      kind: 'paragraph', ordinaryFlow: true,
      flowDomainId: layout.pages[1]!.sectionRegions[0]!.flowDomainIds[0],
    });
  });

  it.each([
    ['LTR left positive', 'left', 12, false, 22],
    ['LTR left negative', 'left', -12, false, -2],
    ['LTR center positive', 'center', 12, false, 92],
    ['LTR center negative', 'center', -12, false, 68],
    ['LTR right positive', 'right', 12, false, 162],
    ['LTR right negative', 'right', -12, false, 138],
    ['RTL left positive', 'left', 12, true, 138],
    ['RTL left negative', 'left', -12, true, 162],
    ['RTL center positive', 'center', 12, true, 68],
    ['RTL center negative', 'center', -12, true, 92],
    ['RTL right positive', 'right', 12, true, -2],
    ['RTL right negative', 'right', -12, true, 22],
  ] as const)(
    'projects parser-owned tblInd placement for %s',
    (_name, justification, indentPt, bidiVisual, expectedXPt) => {
      const section = {
        pageWidth: 200, pageHeight: 100,
        marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
        headerDistance: 5, footerDistance: 5, titlePage: false,
        evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
      } as SectionProps;
      const model = {
        section, body: [ordinaryTable(justification, indentPt, bidiVisual)],
        headers: { default: null, first: null, even: null },
        footers: { default: null, first: null, even: null },
        footnotes: [], endnotes: [], fontFamilyClasses: {},
      } as unknown as DocxDocumentModel;
      const services = createLayoutServices(model, { measureContext: measureContext() });

      const retained = layoutDocument(model, services, { currentDateMs: 0 })
        .pages[0]?.layers.body[0];

      expect(retained?.kind).toBe('table');
      if (retained?.kind !== 'table') throw new Error('expected retained table');
      expect(retained.flowBounds.xPt).toBe(expectedXPt);
      expect(retained.rows[0]?.flowBounds.xPt).toBe(expectedXPt);
    },
  );

  it('retains the note-reference id on the destination-page text placement', () => {
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section, body: [paragraph()],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [{ id: '7', content: [paragraph()] }],
      endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const placement = layout.pages[0]!.layers.body
      .filter((node) => node.kind === 'paragraph')
      .flatMap((node) => node.kind === 'paragraph' ? node.lines : [])
      .flatMap((line) => line.placements)
      .find((candidate) => candidate.kind === 'text');

    expect(placement).toMatchObject({ noteReference: { kind: 'footnote', id: '7' } });
  });

  it('retains a frame paragraph as an out-of-flow placed occurrence', () => {
    const framed = paragraph();
    framed.framePr = {
      dropCap: 'none', lines: 1, wrap: 'around', hAnchor: 'text', vAnchor: 'text',
      hRule: 'auto', hSpace: 0, vSpace: 0, w: 40, h: 20, x: 15, y: 5,
    };
    const following = paragraph();
    (following.runs[0] as { noteRef?: unknown }).noteRef = undefined;
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section, body: [framed, following],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const body = layout.pages[0]!.layers.body;

    const frame = body[0]!;
    const follower = body[1]!;
    expect(frame.kind).toBe('paragraph');
    if (frame.kind !== 'paragraph') throw new Error('expected retained frame paragraph');
    expect(frame.ordinaryFlow).toBe(false);
    expect(frame.advancePt).toBe(0);
    expect(frame.flowBounds.heightPt).toBeCloseTo(
      frame.spacing.beforePt
        + fragmentLineAdvancesPt(frame)
        + frame.spacing.afterPt,
      6,
    );
    expect(frame.flowBounds).toMatchObject({ xPt: 25, yPt: 15 });
    expect(follower.flowBounds.yPt).toBe(10);
  });

  it('retains an effective positioned table without charging ordinary flow', () => {
    const following = paragraph();
    (following.runs[0] as { noteRef?: unknown }).noteRef = undefined;
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section, body: [floatingTable(), following],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const body = layout.pages[0]!.layers.body;

    const floating = body[0]!;
    const follower = body[1]!;
    expect(floating.kind).toBe('table');
    if (floating.kind !== 'table') throw new Error('expected retained floating table');
    expect(floating.ordinaryFlow).toBe(false);
    expect(floating.advancePt).toBeCloseTo(
      floating.rows.reduce((sum, row) => sum + row.advancePt, 0),
      6,
    );
    expect(floating.advancePt).toBeGreaterThan(0);
    expect(floating.flowBounds).toMatchObject({ xPt: 30, yPt: 20 });
    expect(follower.flowBounds.yPt).toBe(10);
  });
});
