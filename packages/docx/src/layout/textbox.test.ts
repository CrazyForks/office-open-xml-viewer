import { describe, expect, it, vi } from 'vitest';
import type { ParagraphLayoutContext } from '../layout-context.js';
import type {
  BodyElement,
  DocParagraph,
  DocRun,
  DocxDocumentModel,
  ShapeRun,
} from '../types.js';
import { createLayoutServices, layoutDocument } from '../renderer.js';
import { acquireShapeTextBoxLayout } from './paragraph.js';
import type { CompleteTextBoxBlockInput } from './textbox-input.js';
import type {
  ParagraphLayout,
  StoryLayout,
  TableLayout,
} from './types.js';

const paragraphContext: ParagraphLayoutContext = {
  lineGrid: { active: false, pitchPt: null },
  characterGrid: { active: false, deltaPt: 0 },
  physicalIndentLeftPt: 0,
  physicalIndentRightPt: 0,
  firstIndentPt: 0,
  lineSpacing: null,
  spaceBeforePt: 0,
  spaceAfterPt: 0,
  baseRtl: false,
  isJustified: false,
  stretchLastLine: false,
  tabStops: [],
  hasRuby: false,
  hasEastAsianText: false,
  kinsoku: { enabled: true, lineStartForbidden: new Set(), lineEndForbidden: new Set() },
  defaultTabPt: 36,
};

function canvas(): CanvasRenderingContext2D {
  return {
    font: '10px sans-serif',
    measureText: () => ({
      width: 10,
      fontBoundingBoxAscent: 8,
      fontBoundingBoxDescent: 2,
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
    }),
  } as unknown as CanvasRenderingContext2D;
}

function paragraph(): ParagraphLayout {
  const bounds = { xPt: 7, yPt: 11, widthPt: 80, heightPt: 10 };
  return {
    kind: 'paragraph',
    id: 'textbox-paragraph',
    source: { story: 'textbox', storyInstance: 'shape:9', path: [0] },
    flowDomainId: 'textbox:shape:9',
    ordinaryFlow: true,
    flowBounds: bounds,
    inkBounds: bounds,
    advancePt: 10,
    spacing: { beforePt: 0, afterPt: 0 },
    contextualSpacing: false,
    lines: [],
    borders: [],
    resources: [],
    drawings: [],
    textBoxes: [],
    events: [],
    exclusions: [],
  };
}

function table(): TableLayout {
  const bounds = { xPt: 7, yPt: 21, widthPt: 80, heightPt: 15 };
  return {
    kind: 'table',
    id: 'textbox-table',
    source: { story: 'textbox', storyInstance: 'shape:9', path: [1] },
    flowDomainId: 'textbox:shape:9',
    ordinaryFlow: true,
    flowBounds: bounds,
    inkBounds: bounds,
    advancePt: 15,
    columnWidthsPt: [80],
    rows: [],
    borders: [],
  };
}

describe('rich text-box story acquisition', () => {
  it('prefers complete parser blocks and retains one nested StoryLayout', () => {
    const richBlocks = [
      { type: 'paragraph', runs: [{ type: 'text', text: 'before' }] },
      {
        type: 'table',
        rows: [{ cells: [{ content: [
          { type: 'paragraph', runs: [{ type: 'text', text: 'cell' }] },
        ] }] }],
      },
      { type: 'unsupportedTextBoxBlock', qName: 'w:altChunk', sourcePath: [2] },
    ] as unknown as readonly CompleteTextBoxBlockInput[];
    const retainedStory: StoryLayout = {
      story: 'textbox',
      flowBounds: { xPt: 7, yPt: 11, widthPt: 80, heightPt: 25 },
      inkBounds: { xPt: 7, yPt: 11, widthPt: 80, heightPt: 25 },
      clipBounds: { xPt: 7, yPt: 11, widthPt: 80, heightPt: 44 },
      blocks: [paragraph(), table()],
      advancePt: 25,
      diagnostics: [{
        code: 'UNSUPPORTED_FEATURE',
        severity: 'warning',
        source: { story: 'textbox', storyInstance: 'shape:9', path: [2] },
        message: 'Unsupported text-box block w:altChunk',
      }],
    };
    const acquireCompleteStory = vi.fn(() => retainedStory);
    const shape = {
      type: 'shape',
      textInsetL: 2,
      textInsetT: 3,
      textInsetR: 4,
      textInsetB: 5,
      textAutofit: 'none',
    } as unknown as ShapeRun;

    const layout = acquireShapeTextBoxLayout(
      shape,
      { xPt: 5, yPt: 8, widthPt: 86, heightPt: 52 },
      {
        id: 'shape:9',
        source: { story: 'textbox', storyInstance: 'shape:9', path: [] },
        flowDomainId: 'shape-owner',
        context: paragraphContext,
        measurer: { context: canvas(), fontFamilyClasses: {} },
        environment: {
          pageIndex: 0,
          totalPages: 1,
          documentHasEastAsianText: false,
        },
        input: {
          kind: 'complete',
          source: { story: 'textbox', storyInstance: 'shape:9', path: [] },
          blocks: richBlocks,
        },
        acquireCompleteStory,
      },
    );

    expect(acquireCompleteStory).toHaveBeenCalledOnce();
    expect(acquireCompleteStory).toHaveBeenCalledWith(expect.objectContaining({
      source: { story: 'textbox', storyInstance: 'shape:9', path: [] },
      blocks: richBlocks,
      container: {
        id: 'shape:9:story',
        kind: 'textbox',
        bounds: { xPt: 7, yPt: 11, widthPt: 80, heightPt: 44 },
        capacity: 'unbounded',
      },
    }));
    expect(layout).toMatchObject({
      kind: 'textbox',
      story: retainedStory,
      flowBounds: { xPt: 5, yPt: 8, widthPt: 86, heightPt: 52 },
      clipBounds: { xPt: 7, yPt: 11, widthPt: 80, heightPt: 44 },
      transform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
    });
    expect(layout).not.toHaveProperty('paragraphs');
  });

  it('uses the production shared story algorithms for paragraph/table order', () => {
    const textParagraph = (text: string): DocParagraph => ({
      type: 'paragraph',
      alignment: 'left',
      indentLeft: 0,
      indentRight: 0,
      indentFirst: 0,
      spaceBefore: 0,
      spaceAfter: 0,
      lineSpacing: null,
      numbering: null,
      tabStops: [],
      runs: [{
        type: 'text',
        text,
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
        fontSize: 10,
        color: null,
        fontFamily: 'serif',
        fontFamilyEastAsia: '',
        isLink: false,
        background: null,
        vertAlign: null,
        hyperlink: null,
      }],
      defaultFontSize: 10,
      defaultFontFamily: 'serif',
      widowControl: false,
    } as unknown as DocParagraph);
    const tableWithContent = (
      content: readonly BodyElement[],
      widthPt: number,
    ): Extract<BodyElement, { type: 'table' }> => ({
      type: 'table',
      colWidths: [widthPt],
      rows: [{
        cells: [{
          content,
          colSpan: 1,
          vMerge: null,
          borders: {
            top: null, bottom: null, left: null, right: null,
            insideH: null, insideV: null,
          },
          background: null,
          vAlign: 'top',
          widthPt,
        }],
        rowHeight: null,
        rowHeightRule: 'auto',
        isHeader: false,
      }],
      borders: {
        top: null, bottom: null, left: null, right: null,
        insideH: null, insideV: null,
      },
      cellMarginTop: 0,
      cellMarginBottom: 0,
      cellMarginLeft: 0,
      cellMarginRight: 0,
      jc: 'left',
      layout: 'fixed',
      overlap: 'overlap',
      tblpPr: null,
    } as unknown as Extract<BodyElement, { type: 'table' }>);
    const nestedTable = tableWithContent([
      textParagraph('nested cell') as unknown as BodyElement,
    ], 40);
    const tableBlock = tableWithContent([
      textParagraph('cell before nested') as unknown as BodyElement,
      nestedTable,
      textParagraph('cell after nested') as unknown as BodyElement,
    ], 60);
    const floatingParagraph = textParagraph('before');
    floatingParagraph.runs.push({
      type: 'shape',
      widthPt: 18,
      heightPt: 12,
      anchorXPt: 35,
      anchorYPt: 0,
      anchorXFromMargin: false,
      anchorYFromPara: true,
      wrapMode: 'square',
      zOrder: 2,
      subpaths: [],
      presetGeometry: 'rect',
      fill: 'D9EAF7',
      stroke: null,
    } as unknown as DocRun);
    const shape = {
      type: 'shape',
      widthPt: 100,
      heightPt: 20,
      anchorXPt: 0,
      anchorYPt: 0,
      anchorXFromMargin: false,
      anchorYFromPara: true,
      zOrder: 0,
      subpaths: [],
      presetGeometry: 'rect',
      fill: null,
      stroke: null,
      wrapMode: 'none',
      textInsetL: 2,
      textInsetT: 2,
      textInsetR: 2,
      textInsetB: 2,
      textAutofit: 'sp',
      textBlocks: [{ text: 'lossy fallback', fontSizePt: 10 }],
      textBoxContent: [
        floatingParagraph,
        tableBlock,
        {
          type: 'unsupportedTextBoxBlock',
          qName: 'w:altChunk',
          sourcePath: [2],
        },
        textParagraph('after'),
      ],
    } as unknown as ShapeRun;
    const verticalTableBlock = tableWithContent([
      textParagraph('表の前') as unknown as BodyElement,
      tableWithContent([
        textParagraph('内側') as unknown as BodyElement,
      ], 35),
      textParagraph('表の後') as unknown as BodyElement,
    ], 55);
    const verticalShape = {
      ...shape,
      anchorXPt: 110,
      textVert: 'eaVert',
      textBlocks: [{ text: 'lossy vertical fallback', fontSizePt: 10 }],
      textBoxContent: [
        textParagraph('縦書き'),
        verticalTableBlock,
      ],
    } as unknown as ShapeRun;
    const host = textParagraph('');
    host.runs = [
      shape as unknown as DocRun,
      verticalShape as unknown as DocRun,
    ];
    const model = {
      section: {
        pageWidth: 240,
        pageHeight: 160,
        marginTop: 10,
        marginRight: 10,
        marginBottom: 10,
        marginLeft: 10,
        headerDistance: 5,
        footerDistance: 5,
        titlePage: false,
        evenAndOddHeaders: false,
        sectionStart: 'nextPage',
        columns: null,
      },
      body: [host],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [],
      endnotes: [],
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;

    const layout = layoutDocument(
      model,
      createLayoutServices(model, { measureContext: canvas() }),
      { currentDateMs: 0 },
    );
    const bodyParagraph = layout.pages[0]?.layers.body.find(
      (node) => node.kind === 'paragraph',
    );
    if (!bodyParagraph || bodyParagraph.kind !== 'paragraph') {
      throw new Error('expected retained body paragraph');
    }
    const textBox = bodyParagraph.textBoxes[0];

    expect(textBox?.flowBounds.heightPt).toBeGreaterThan(20);
    expect(textBox?.clipBounds).toBeUndefined();
    expect(textBox?.story.blocks.map((block) => block.kind)).toEqual([
      'paragraph', 'table', 'paragraph',
    ]);
    expect(textBox?.story.blocks.map((block) => block.source.path)).toEqual([
      [0], [1], [3],
    ]);
    expect(textBox?.story.diagnostics).toEqual([{
      code: 'UNSUPPORTED_FEATURE',
      severity: 'warning',
      source: expect.objectContaining({ story: 'textbox', path: [2] }),
      message: 'Unsupported text-box block w:altChunk',
    }]);
    expect(layout.diagnostics).toEqual(textBox?.story.diagnostics);
    expect(textBox?.story.blocks[0]).toMatchObject({
      kind: 'paragraph',
      lines: [expect.objectContaining({
        placements: expect.arrayContaining([
          expect.objectContaining({ kind: 'text', text: 'before' }),
        ]),
      })],
    });
    const retainedFloatingParagraph = textBox?.story.blocks[0];
    expect(retainedFloatingParagraph?.kind).toBe('paragraph');
    if (retainedFloatingParagraph?.kind !== 'paragraph') {
      throw new Error('expected retained floating-content paragraph');
    }
    expect(retainedFloatingParagraph.drawings).toHaveLength(1);
    expect(retainedFloatingParagraph.exclusions).toEqual([
      expect.objectContaining({ wrap: 'square' }),
    ]);
    const retainedTable = textBox?.story.blocks[1];
    expect(retainedTable?.kind).toBe('table');
    if (retainedTable?.kind !== 'table') throw new Error('expected retained text-box table');
    expect(retainedTable.rows[0]?.cells[0]?.blocks.map((block) => block.layout.kind))
      .toEqual(['paragraph', 'table', 'paragraph']);

    const verticalTextBox = bodyParagraph.textBoxes[1];
    expect(verticalTextBox).toMatchObject({
      verticalMode: 'eaVert',
      transform: { a: 0, b: 1, c: -1, d: 0 },
    });
    expect(verticalTextBox?.story.blocks.map((block) => block.kind))
      .toEqual(['paragraph', 'table']);
    const verticalParagraph = verticalTextBox?.story.blocks[0];
    expect(verticalParagraph?.kind).toBe('paragraph');
    const topLevelOrientations = verticalParagraph?.kind === 'paragraph'
      ? verticalParagraph.lines.flatMap((line) => line.placements)
          .flatMap((placement) => placement.kind === 'text'
            ? placement.paintOps.map((operation) => operation.glyphOrientation)
            : [])
      : [];
    expect(topLevelOrientations).toContain('upright');
    const verticalTable = verticalTextBox?.story.blocks[1];
    expect(verticalTable?.kind).toBe('table');
    const cellParagraph = verticalTable?.kind === 'table'
      ? verticalTable.rows[0]?.cells[0]?.blocks[0]?.layout
      : undefined;
    expect(cellParagraph?.kind).toBe('paragraph');
    const cellOrientations = cellParagraph?.kind === 'paragraph'
      ? cellParagraph.lines.flatMap((line) => line.placements)
          .flatMap((placement) => placement.kind === 'text'
            ? placement.paintOps.map((operation) => operation.glyphOrientation)
            : [])
      : [];
    expect(cellOrientations).toContain('upright');
  });
});
