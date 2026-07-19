import { canvasFontString, PT_TO_PX } from '@silurus/ooxml-core';
import type { DocxTextRunInfo } from '../renderer.js';
import { selectDocumentLayoutPage } from './document-layout-variants.js';
import {
  composeAffine,
  cssTransformFor,
  mapAffinePoint,
  scaleAffine,
  translationAffine,
} from '../paint/affine.js';
import type {
  DocumentLayout,
  DrawingLayout,
  LayoutPage,
  LayoutServices,
  Matrix2DData,
  PageLayerRoot,
  PagePaintDrawingEntry,
  PaintNode,
  ParagraphLayout,
  PointPt,
  TableLayout,
  TextBoxLayout,
  TextPlacement,
} from './types.js';

export interface TextRunsForPageOptions {
  readonly scale: number;
}

export interface SelectedTextRunsForPageOptions {
  readonly defaultCurrentDateMs: number;
  readonly currentDate?: Date | number;
  readonly width?: number;
}

interface ProjectionContext {
  readonly drawingEntries: ReadonlyMap<string, PagePaintDrawingEntry>;
  readonly rootPointToCss: ReadonlyMap<string, Matrix2DData>;
  readonly emittedTextBoxes: Set<string>;
  readonly runs: DocxTextRunInfo[];
}

interface NodeProjection {
  readonly pointToCss: Matrix2DData;
  readonly layoutTranslationPt: PointPt;
  readonly rootNodeId: string;
}

function pageRegionsByDomain(
  page: LayoutPage,
): ReadonlyMap<string, LayoutPage['sectionRegions'][number]> {
  const byId = new Map(page.sectionRegions.map((region) => [region.id, region]));
  const byDomain = new Map<string, LayoutPage['sectionRegions'][number]>();
  for (const region of page.sectionRegions) {
    for (const flowDomainId of region.flowDomainIds) {
      byDomain.set(flowDomainId, region);
    }
  }
  for (const domain of page.flowDomains) {
    if (
      (domain.kind === 'footnote' || domain.kind === 'endnote')
      && domain.sectionRegionId
    ) {
      const region = byId.get(domain.sectionRegionId);
      if (region) byDomain.set(domain.id, region);
    }
  }
  return byDomain;
}

function pointToCssForRoot(
  regionByDomain: ReadonlyMap<string, LayoutPage['sectionRegions'][number]>,
  root: Pick<PageLayerRoot, 'coordinateSpace' | 'node'>,
  scale: number,
): Matrix2DData {
  if (root.coordinateSpace === 'upright-physical') return scaleAffine(scale);
  const matrix = regionByDomain.get(root.node.flowDomainId)
    ?.coordinateSpace.logicalToPhysical;
  return matrix
    ? composeAffine(scaleAffine(scale), matrix)
    : scaleAffine(scale);
}

function pointToCssForEntry(
  context: ProjectionContext,
  entry: PagePaintDrawingEntry,
): Matrix2DData {
  const rootPointToCss = context.rootPointToCss.get(entry.rootNodeId);
  if (!rootPointToCss) {
    throw new Error(`Drawing entry ${entry.node.id} references missing root ${entry.rootNodeId}`);
  }
  let pointToCss = rootPointToCss;
  for (const frame of entry.frames) {
    if (frame.kind === 'transform') {
      pointToCss = composeAffine(pointToCss, frame.transform);
    }
  }
  return pointToCss;
}

function projectTextPlacement(
  placement: TextPlacement,
  pointToCss: Matrix2DData,
): DocxTextRunInfo {
  const origin = mapAffinePoint(pointToCss, placement.bounds);
  const inlineScale = Math.hypot(pointToCss.a, pointToCss.b);
  const blockScale = Math.hypot(pointToCss.c, pointToCss.d);
  const transform = cssTransformFor(pointToCss);
  const letterSpacingPt = placement.paintOps[0]?.letterSpacingPt ?? 0;
  return {
    text: placement.text,
    x: origin.xPt,
    y: origin.yPt,
    w: placement.bounds.widthPt * inlineScale,
    h: placement.bounds.heightPt * blockScale,
    fontSize: placement.fontSizePt * blockScale,
    font: canvasFontString(
      placement.fontRoute,
      placement.fontSizePt * blockScale,
      placement.fontWeight,
      placement.fontStyle,
    ),
    ...(letterSpacingPt !== 0
      ? { letterSpacingPx: letterSpacingPt * inlineScale }
      : {}),
    ...(transform ? { transform } : {}),
    ...(placement.hyperlink ? { hyperlink: placement.hyperlink } : {}),
    ...(placement.tateChuYoko ? { eastAsianVert: true } : {}),
  };
}

function placedChildProjection(
  child: ParagraphLayout | TableLayout,
  placement: Readonly<{ xPt: number; yPt: number }>,
  parent: NodeProjection,
): NodeProjection {
  const dxPt = placement.xPt - child.flowBounds.xPt;
  const dyPt = placement.yPt - child.flowBounds.yPt;
  return {
    pointToCss: composeAffine(
      parent.pointToCss,
      translationAffine(dxPt, dyPt),
    ),
    layoutTranslationPt: {
      xPt: parent.layoutTranslationPt.xPt + dxPt,
      yPt: parent.layoutTranslationPt.yPt + dyPt,
    },
    rootNodeId: parent.rootNodeId,
  };
}

function drawingTextBoxes(
  byId: ReadonlyMap<string, TextBoxLayout>,
  drawing: DrawingLayout,
): readonly TextBoxLayout[] {
  return (drawing.textBoxIds ?? []).flatMap((id) => {
    const textBox = byId.get(id);
    return textBox ? [textBox] : [];
  });
}

function visitTextBox(
  textBox: TextBoxLayout,
  projection: NodeProjection,
  context: ProjectionContext,
): void {
  if (context.emittedTextBoxes.has(textBox.id)) return;
  context.emittedTextBoxes.add(textBox.id);
  const textBoxProjection: NodeProjection = {
    ...projection,
    pointToCss: composeAffine(projection.pointToCss, textBox.transform),
  };
  for (const block of textBox.story.blocks) {
    visitNode(block, textBoxProjection, context);
  }
}

function visitDrawingTextBoxes(
  textBoxesById: ReadonlyMap<string, TextBoxLayout>,
  drawing: DrawingLayout,
  projection: NodeProjection,
  context: ProjectionContext,
): void {
  const textBoxes = drawingTextBoxes(textBoxesById, drawing);
  if (textBoxes.length === 0) return;
  const retainedEntry = context.drawingEntries.get(drawing.id);
  let drawingProjection = projection;
  if (
    retainedEntry
    && retainedEntry.rootNodeId === projection.rootNodeId
  ) {
    drawingProjection = {
      pointToCss: pointToCssForEntry(context, retainedEntry),
      layoutTranslationPt: retainedEntry.layoutTranslationPt,
      rootNodeId: retainedEntry.rootNodeId,
    };
  }
  const translation = drawingProjection.layoutTranslationPt;
  const undoX = drawing.anchorLayer?.horizontalOwnership === 'page'
    ? -translation.xPt : 0;
  const undoY = drawing.anchorLayer?.verticalOwnership === 'page'
    ? -translation.yPt : 0;
  const ownedProjection = undoX === 0 && undoY === 0
    ? drawingProjection
    : {
        ...drawingProjection,
        pointToCss: composeAffine(
          drawingProjection.pointToCss,
          translationAffine(undoX, undoY),
        ),
      };
  for (const textBox of textBoxes) {
    visitTextBox(textBox, ownedProjection, context);
  }
}

function visitParagraph(
  paragraph: ParagraphLayout,
  projection: NodeProjection,
  context: ProjectionContext,
): void {
  for (const line of paragraph.lines) {
    for (const placement of line.placements) {
      if (placement.kind === 'text') {
        context.runs.push(projectTextPlacement(placement, projection.pointToCss));
      }
    }
  }

  const indexedDrawings = paragraph.drawings.map((drawing, index) => ({
    drawing,
    index,
  }));
  indexedDrawings.sort((left, right) => (
    (left.drawing.anchorLayer?.sourceOrder ?? left.index)
      - (right.drawing.anchorLayer?.sourceOrder ?? right.index)
    || left.index - right.index
  ));
  const textBoxesById = new Map(
    paragraph.textBoxes.map((textBox) => [textBox.id, textBox]),
  );
  const ownedTextBoxIds = new Set<string>();
  for (const { drawing } of indexedDrawings) {
    for (const id of drawing.textBoxIds ?? []) ownedTextBoxIds.add(id);
    visitDrawingTextBoxes(textBoxesById, drawing, projection, context);
  }
  for (const textBox of paragraph.textBoxes) {
    if (!ownedTextBoxIds.has(textBox.id)) {
      visitTextBox(textBox, projection, context);
    }
  }
}

function visitTable(
  table: TableLayout,
  projection: NodeProjection,
  context: ProjectionContext,
): void {
  for (const row of table.rows) {
    for (const cell of row.cells) {
      const ownsContinuationPaint = 'visualMergeOwnership' in cell
        && cell.visualMergeOwnership === 'continuation';
      if (cell.verticalMerge === 'continue' && !ownsContinuationPaint) continue;
      for (const block of cell.blocks) {
        const child = block.layout;
        visitNode(child, placedChildProjection(child, {
          xPt: cell.contentBounds.xPt
            + (child.kind === 'table' ? child.flowBounds.xPt : 0),
          yPt: cell.flowBounds.yPt + block.offsetPt
            + (child.kind === 'table' ? child.flowBounds.yPt : 0),
        }, projection), context);
      }
    }
  }
  for (const placement of table.resolvedFloatingTables ?? []) {
    visitNode(placement.child, placedChildProjection(placement.child, {
      xPt: placement.xPt - projection.layoutTranslationPt.xPt,
      yPt: placement.yPt - projection.layoutTranslationPt.yPt,
    }, projection), context);
  }
}

function visitNode(
  node: PaintNode,
  projection: NodeProjection,
  context: ProjectionContext,
): void {
  switch (node.kind) {
    case 'paragraph':
      visitParagraph(node, projection, context);
      return;
    case 'table':
      visitTable(node, projection, context);
      return;
    case 'note':
      for (const block of node.story.blocks) visitNode(block, projection, context);
      return;
    case 'textbox':
      visitTextBox(node, projection, context);
      return;
    case 'drawing': {
      const entry = context.drawingEntries.get(node.id);
      for (const textBox of entry?.textBoxes ?? []) {
        visitTextBox(textBox, projection, context);
      }
      return;
    }
    default: {
      const exhaustive: never = node;
      throw new Error(`Unknown text-index node: ${String(exhaustive)}`);
    }
  }
}

/**
 * Projects retained point-space text placements into the existing public
 * selection/search run shape. Sequence follows semantic reading order; paint
 * order contributes only already-materialized anchor frame geometry.
 */
export function textRunsForPage(
  layout: DocumentLayout,
  pageIndex: number,
  options: TextRunsForPageOptions,
): DocxTextRunInfo[] {
  if (!Number.isFinite(options.scale) || options.scale <= 0) {
    throw new RangeError(`Text projection scale must be positive: ${options.scale}`);
  }
  const page = layout.pages[pageIndex];
  if (!page) throw new RangeError(`Page index ${pageIndex} is out of range`);
  const roots = new Map(page.layers.roots.map((root) => [root.node.id, root]));
  const regionByDomain = pageRegionsByDomain(page);
  const rootPointToCss = new Map(page.layers.roots.map((root) => [
    root.node.id,
    pointToCssForRoot(regionByDomain, root, options.scale),
  ]));
  const drawingEntries = new Map<string, PagePaintDrawingEntry>();
  for (const entry of page.layers.paintOrder) {
    if (entry.kind === 'drawing') drawingEntries.set(entry.node.id, entry);
  }
  const context: ProjectionContext = {
    drawingEntries,
    rootPointToCss,
    emittedTextBoxes: new Set(),
    runs: [],
  };
  for (const nodeId of page.readingOrder) {
    const root = roots.get(nodeId);
    if (!root) throw new Error(`Reading-order node ${nodeId} is not a page root`);
    const pointToCss = rootPointToCss.get(nodeId);
    if (!pointToCss) throw new Error(`Reading-order node ${nodeId} has no page projection`);
    visitNode(root.node, {
      pointToCss,
      layoutTranslationPt: { xPt: 0, yPt: 0 },
      rootNodeId: root.node.id,
    }, context);
  }
  return context.runs;
}

/** Select the same keyed layout variant as paint, then project its retained text. */
export function textRunsForSelectedPage(
  services: LayoutServices,
  pageIndex: number,
  options: SelectedTextRunsForPageOptions,
): DocxTextRunInfo[] {
  const selected = selectDocumentLayoutPage(services, {
    currentDate: options.currentDate,
    defaultCurrentDateMs: options.defaultCurrentDateMs,
  }, pageIndex);
  const scale = (
    options.width ?? selected.page.geometry.widthPt * PT_TO_PX
  ) / selected.page.geometry.widthPt;
  return textRunsForPage(selected.layout, pageIndex, { scale });
}
