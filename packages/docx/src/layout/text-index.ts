import {
  composeAffine,
  translationAffine,
} from './affine.js';
import type {
  DocumentLayout,
  DrawingLayout,
  LayoutPage,
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

export interface TextRunGeometry {
  readonly placement: TextPlacement;
  readonly pointToPage: Matrix2DData;
  /** Source `w14:paraId` of the owning paragraph, when authored. */
  readonly paragraphId?: string;
}

interface ProjectionContext {
  readonly drawingEntries: ReadonlyMap<string, PagePaintDrawingEntry>;
  readonly rootPointToPage: ReadonlyMap<string, Matrix2DData>;
  readonly emittedTextBoxes: Set<string>;
  readonly runs: TextRunGeometry[];
}

interface NodeProjection {
  readonly pointToPage: Matrix2DData;
  readonly layoutTranslationPt: PointPt;
  readonly rootNodeId: string;
}

const IDENTITY_AFFINE = Object.freeze({
  a: 1, b: 0, c: 0, d: 1, e: 0, f: 0,
}) satisfies Matrix2DData;

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
    if (domain.kind !== 'footnote' && domain.kind !== 'endnote') continue;
    const storyRegion = domain.sectionRegionId
      ? byId.get(domain.sectionRegionId)
      : page.sectionRegions[0];
    if (!storyRegion) {
      throw new Error(
        `${domain.id} references missing page story region ${domain.sectionRegionId ?? '<default>'}`,
      );
    }
    byDomain.set(domain.id, storyRegion);
  }
  return byDomain;
}

function pointToPageForRoot(
  regionByDomain: ReadonlyMap<string, LayoutPage['sectionRegions'][number]>,
  root: Pick<PageLayerRoot, 'coordinateSpace' | 'node'>,
): Matrix2DData {
  if (root.coordinateSpace === 'upright-physical') return IDENTITY_AFFINE;
  const matrix = regionByDomain.get(root.node.flowDomainId)
    ?.coordinateSpace.logicalToPhysical;
  return matrix ?? IDENTITY_AFFINE;
}

function pointToPageForEntry(
  context: ProjectionContext,
  entry: PagePaintDrawingEntry,
): Matrix2DData {
  const rootPointToPage = context.rootPointToPage.get(entry.rootNodeId);
  if (!rootPointToPage) {
    throw new Error(`Drawing entry ${entry.node.id} references missing root ${entry.rootNodeId}`);
  }
  let pointToPage = rootPointToPage;
  for (const frame of entry.frames) {
    if (frame.kind === 'transform') {
      pointToPage = composeAffine(pointToPage, frame.transform);
    }
    // Clip frames affect visible pixels, but the former post-paint callback
    // reported retained text geometry even when a run was partially clipped.
  }
  return pointToPage;
}

function placedChildProjection(
  child: ParagraphLayout | TableLayout,
  placement: Readonly<{ xPt: number; yPt: number }>,
  parent: NodeProjection,
): NodeProjection {
  const dxPt = placement.xPt - child.flowBounds.xPt;
  const dyPt = placement.yPt - child.flowBounds.yPt;
  return {
    pointToPage: composeAffine(
      parent.pointToPage,
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
    pointToPage: composeAffine(projection.pointToPage, textBox.transform),
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
  const ownedProjection = drawingOwnedContentProjection(drawing, projection, context);
  for (const textBox of textBoxes) {
    visitTextBox(textBox, ownedProjection, context);
  }
}

function drawingOwnedContentProjection(
  drawing: DrawingLayout,
  projection: NodeProjection,
  context: ProjectionContext,
): NodeProjection {
  const retainedEntry = context.drawingEntries.get(drawing.id);
  let drawingProjection = projection;
  if (
    retainedEntry
    && retainedEntry.rootNodeId === projection.rootNodeId
  ) {
    drawingProjection = {
      pointToPage: pointToPageForEntry(context, retainedEntry),
      layoutTranslationPt: retainedEntry.layoutTranslationPt,
      rootNodeId: retainedEntry.rootNodeId,
    };
  }
  const translation = drawingProjection.layoutTranslationPt;
  const undoX = drawing.anchorLayer?.horizontalOwnership === 'page'
    ? -translation.xPt : 0;
  const undoY = drawing.anchorLayer?.verticalOwnership === 'page'
    ? -translation.yPt : 0;
  let ownedProjection = undoX === 0 && undoY === 0
    ? drawingProjection
    : {
        ...drawingProjection,
        pointToPage: composeAffine(
          drawingProjection.pointToPage,
          translationAffine(undoX, undoY),
        ),
      };
  if (drawing.orientation === 'upright-physical') {
    if (!drawing.transform) {
      throw new Error(`Upright physical drawing ${drawing.id} is missing its logical transform`);
    }
    ownedProjection = {
      ...ownedProjection,
      pointToPage: composeAffine(ownedProjection.pointToPage, drawing.transform),
    };
  }
  return ownedProjection;
}

function visitParagraph(
  paragraph: ParagraphLayout,
  projection: NodeProjection,
  context: ProjectionContext,
): void {
  for (const line of paragraph.lines) {
    for (const placement of line.placements) {
      if (placement.kind === 'text') {
        context.runs.push(Object.freeze({
          placement,
          pointToPage: projection.pointToPage,
          ...(paragraph.paragraphId !== undefined
            ? { paragraphId: paragraph.paragraphId }
            : {}),
        }));
      }
    }
  }

  const textBoxesById = new Map(
    paragraph.textBoxes.map((textBox) => [textBox.id, textBox]),
  );
  const ownedTextBoxIds = new Set<string>();
  const drawingsInSourceOrder = paragraph.drawings
    .map((drawing, index) => {
      const runIndex = drawing.source.path.at(-1);
      if (runIndex === undefined || !Number.isSafeInteger(runIndex) || runIndex < 0) {
        throw new Error(`Drawing ${drawing.id} has no retained paragraph run index`);
      }
      return { drawing, index, runIndex };
    })
    .sort((left, right) => left.runIndex - right.runIndex || left.index - right.index);
  // Both anchored and inline drawings retain their paragraph run index as the
  // terminal SourceRef path component. That is the one comparable source-order
  // domain; anchor stacking ordinals and drawings-array indexes are not mixed.
  for (const { drawing } of drawingsInSourceOrder) {
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
      const ownedProjection = drawingOwnedContentProjection(node, projection, context);
      for (const textBox of entry?.textBoxes ?? []) {
        visitTextBox(textBox, ownedProjection, context);
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
 * Indexes retained text placements in physical page points. Sequence follows
 * semantic reading order; paint order contributes only already-materialized
 * anchor frame geometry.
 */
export function textRunGeometryForPage(
  layout: DocumentLayout,
  pageIndex: number,
): readonly TextRunGeometry[] {
  const page = layout.pages[pageIndex];
  if (!page) throw new RangeError(`Page index ${pageIndex} is out of range`);
  const roots = new Map(page.layers.roots.map((root) => [root.node.id, root]));
  const regionByDomain = pageRegionsByDomain(page);
  const rootPointToPage = new Map(page.layers.roots.map((root) => [
    root.node.id,
    pointToPageForRoot(regionByDomain, root),
  ]));
  const drawingEntries = new Map<string, PagePaintDrawingEntry>();
  for (const entry of page.layers.paintOrder) {
    if (entry.kind === 'drawing') drawingEntries.set(entry.node.id, entry);
  }
  const context: ProjectionContext = {
    drawingEntries,
    rootPointToPage,
    emittedTextBoxes: new Set(),
    runs: [],
  };
  for (const nodeId of page.readingOrder) {
    const root = roots.get(nodeId);
    if (!root) throw new Error(`Reading-order node ${nodeId} is not a page root`);
    const pointToPage = rootPointToPage.get(nodeId);
    if (!pointToPage) throw new Error(`Reading-order node ${nodeId} has no page projection`);
    visitNode(root.node, {
      pointToPage,
      layoutTranslationPt: { xPt: 0, yPt: 0 },
      rootNodeId: root.node.id,
    }, context);
  }
  return Object.freeze(context.runs);
}
