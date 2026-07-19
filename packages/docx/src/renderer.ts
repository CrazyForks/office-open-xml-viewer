import type {
  DocxDocumentModel, BodyElement, DocParagraph, DocTable, DocTableRow, DocTableCell, CellElement,
  DocRun, DocxTextRun, ImageRun, ChartRun, ShapeRun, ShapeFill, TextPath, ShapeText, FieldRun, HeaderFooter, HeadersFooters, BorderSpec, TableBorders, CellBorders,
  TabStop, ParagraphBorders, ParaBorderEdge, DocxRunBorder, SectionProps, SectionGeom, PageNumType, PageBorders, PageBorderEdge, DocNote, NumberingInfo, ColumnsSpec, FramePr, TblpPr, DocSettings,
} from './types';
import { docxRenderedFontFamilies } from './document-content.js';
import type { ArrowEnd, Stroke } from '@silurus/ooxml-core';
import { textRunPaintInfo } from './paint/text-run-info.js';
import {
  buildCustomPath,
  buildShapePath,
  renderPresetShape,
  hasPreset,
  getCachedSvgImageByPath,
  preferVectorBlip,
  hexToRgba,
  autoContrastColor,
  resolveFill,
  applyStroke,
  drawArrowHead,
  lineEndRetract,
  retractLineEndpoint,
  getConnectorAnchors,
  mathToMathML,
  recolorSvg,
  crispOffset,
  PT_TO_PX,
  resolveBaseDirection,
  isHTMLCanvas,
  defaultDpr,
  clampCanvasSize,
  classifyCjkFont,
  cjkFallbackChain,
  NON_CJK_SANS_FALLBACKS,
  NON_CJK_SERIF_FALLBACKS,
  DEFAULT_KINSOKU_RULES,
  kinsokuAdjustedSplit,
  crossRunKinsokuRetract,
  isCjkBreakChar,
  classifyFontGeneric,
  isComplexScriptCodePoint,
  getCachedBitmapByPath,
  dropBitmapCacheByPath,
  acquireBitmapCacheLease,
  deferBitmapCloseWhileLeased,
  applyDuotone,
  imageNaturalSize,
  drawImageCropped,
  metafileRasterSize,
  symbolFontToUnicode,
  isSymbolFontFamily,
  symbolTextToUnicodeSegments,
  docxBorderDashArray,
  fillDoubleBorder,
  drawUnderline,
  renderChart,
  graphemeClusterOffsets,
} from '@silurus/ooxml-core';
import type { CanvasFontRoute, MathNode, MathRenderer, KinsokuRules, HyperlinkTarget, NumberFormat, Duotone, ResolvedLocalFontMetric } from '@silurus/ooxml-core';
import { docxUnderlineToDrawingML } from './underline-map.js';
import { intendedSingleLinePx, correctLineMetrics } from './font-metrics.js';
import {
  resolveBorderConflict,
  resolveCellEdges,
  type CellEdgeFlags,
  type ResolvedCellEdges,
} from './cell-border-conflict.js';
import {
  segmentsHaveRtl,
  computeLineVisualOrder,
  resolveAlignEdge,
  jcIsFullyJustified,
  jcStretchesLastLine,
  type AlignEdge,
  type LineVisualOrder,
} from './bidi-line.js';
import {
  type FloatRect,
  FLOAT_OVERLAP_EPS,
  isWrapFloat,
  resolveLineFloatWindow,
  skipPastTopAndBottom,
} from './float-layout.js';
import {
  distributeLineSlack,
  distributedDelta,
  shrinkFitCompression,
  type SegStretch,
} from './text-distribute.js';
import {
  computeKashidaDistribution,
  type KashidaLevel,
  type KashidaSegmentPlan,
} from './kashida-justify.js';
import {
  type FrameBox,
  computeFrameBox,
  frameXContainer,
  registerFrameFloat,
  pushFloatRect,
} from './frame-geometry.js';
import {
  type FloatTableBox,
  computeFloatTableBox,
  registerTableFloat,
  floatTableWrapSide,
  resolveFloatingTableBoxPt,
} from './float-table-geometry.js';
import {
  xContainer,
  yContainer,
  resolveAnchorX,
  resolveAnchorY,
} from './anchor-geometry.js';
import {
  findMergeEndRow,
  rowGridBefore,
  applyTableRowBoundaryFootprints,
  resolveTableRowContentHeights,
  resolveTableRowHeights,
  resolveSingleRowHeight,
} from './table-geometry.js';
import { selectLargestFittingEnd } from './line-fit-policy.js';
import {
  computeSectionColumns as computeColumns,
  enterTableCellStoryContext,
  resolveDocumentLayoutSettings,
  resolveParagraphLayoutContext,
  resolveSectionLayoutContext,
  toLegacyDocGridContext,
  type DocumentLayoutSettings,
  type ParagraphLayoutContext,
  type SectionLayoutContext,
  type StoryContext,
} from './layout-context.js';
import { canvasFontString, justifiedPiecePositions } from '@silurus/ooxml-core';
import type {
  BlockLayoutAlgorithms,
  BodyFlowRegistryDeltaPt,
  BodyFlowRegistrySnapshotPt,
  DrawingMLCollisionRegistrySnapshotPt,
  FlowBlockInput,
  LayoutServices,
  FloatRegistryEntryPt,
  FloatRegistrySnapshotPt,
  FloatingTablePlacementLayout,
  FloatRegistryDeltaPt,
  Matrix2DData,
  DrawingMLCollisionEntryPt,
  NoteLayout,
  ParagraphLayout,
  SourceRef,
  StoryBlockInput,
  StoryLayout,
  TableLayout,
  TableLayoutInput,
} from './layout/types.js';
import type { CompleteTextBoxBlockInput } from './layout/textbox-input.js';
import type { DocumentLayout as RetainedDocumentLayout } from './layout/types.js';
import { normalizeLayoutOptions } from './layout/options.js';
import {
  beginFloatingTablePlacementTransaction,
  floatingTableRegistryDelta,
  resolveFloatingTablePlacementInTransaction,
  validateFloatingTableRegistryDelta,
} from './layout/floating-table-transaction.js';
import type { LayoutOptions } from './layout/options.js';
import {
  paginatedFlowHasPaginationDependentFields,
  paginationFieldGeometryFingerprint,
  paginationFieldFlowGeometry,
  resolvePaginationFieldLayout,
} from './layout/pagination-fields.js';
import {
  createCanvasPaintResourcePainter,
  paintLayoutPageContent,
  paintLayoutPage as paintRetainedLayoutPage,
} from './paint/canvas-page.js';
import { canonicalCanvasPaintResourceHandlers } from './paint/canonical-resource-handlers.js';
import { paintPlacedParagraphLayout, paintTextBoxLayout } from './paint/canvas-text.js';
import { paintDrawingLayout } from './paint/canvas-drawing.js';
import type { CanvasPaintResourcePainter } from './paint/types.js';
import { createDocumentPaintResourceRegistry } from './layout/production-paint-resources.js';
import {
  createProductionPaintResourceSession,
  unavailablePaintResourceHandle,
} from './paint/resource-session.js';
import {
  mathResourceKey,
  bodyMathOccurrences,
  createImageMetadataService,
  createMathMetadataService,
  documentMathOccurrences,
  documentImageMetadataRecords,
  type MathLayoutResource,
} from './layout/resources.js';
import { createFontResolver, type FontInventoryFace } from './layout/font-service.js';
import {
  attachBodyLayoutKernel,
  attachPrivateResourceLookup,
  attachPaintResourceRegistry,
  createFieldAcquisitionServicesView,
  createLayoutServicesRuntimeView,
  fieldAcquisitionContextOf,
  layoutVariantStoreOf,
  paintResourceRegistryOf,
  privateResourceLookupOf,
  type PageFieldAcquisitionContext,
} from './layout/runtime-state.js';
import {
  attachStoryBlockLayoutAlgorithms,
  layoutStory as layoutSharedStory,
} from './layout/stories.js';
import {
  attachDocumentLayoutVariants,
  selectDocumentLayoutPage,
} from './layout/document-layout-variants.js';
import {
  footnoteIdsInRetainedLines,
  footnoteIdsInRetainedSlice,
  noteReferenceIdsInDocumentOrder,
} from './layout/note-reference-ownership.js';
import { isFirstSectionOwnedPage } from './layout/section-page-identity.js';
import type {
  BodyAcquisitionLocation,
  BodyLayoutKernel,
  BodyLayoutSession,
  PageAnchorPrescanInput,
  BodyParagraphAcquisitionInput,
  BodyTableAcquisitionInput,
} from './layout/body-layout-kernel.js';
import { NoteCapacityExceededError } from './layout/body-layout-kernel.js';
import { FlowCapacityExceededError } from './layout/flow.js';
import { createBodyLayoutInput } from './body-layout-input.js';
import { paginateBody } from './layout/body-paginator.js';
import { projectBodyOccurrence } from './layout/occurrence-projection.js';
import {
  calcEffectiveFontPx,
  createTextLayoutService,
  classifyDocxFontGeneric,
  EAST_ASIAN_RE,
  nextTabStop,
  shapeRunToDocRun,
  snapshotLocalMetrics,
  type GlyphMeasureRequest,
} from './layout/text.js';
import { DOCX_GOOGLE_FONTS, docxFontPreloadNames } from './google-fonts.js';
import {
  effectiveTablePositioning,
  anchorAcquisitionInput,
  internalDocumentModel,
  numberingMarkerShapeInput,
  paragraphAcquisitionInput,
  paragraphMarkShapeInput,
  publicAnchorBridge,
  tableParticipatesInOrdinaryFlow,
  textBoxAcquisitionInput,
} from './parser-model.js';
import {
  normalizeInternalDocumentModel,
  bodySectionIndexInput,
  sectionPlacementInputFromBody,
} from './parser-model.js';
import {
  logicalSectionGeometry as logicalGeomOf,
  physicalSectionGeometry as physicalGeomOf,
  sectionBodyInsetPt as bodyMarginInsetPt,
  createBodySectionIndex,
} from './layout/context.js';
import {
  applyNumberingBodyOffset,
  resolveNumberingMarkerGeometry,
  shapeNumberingMarkerText,
  type NumberingMarkerTextLayout,
} from './layout/numbering-marker.js';
import { paintNumberingMarkerText } from './paint/numbering-marker.js';
import { tableColumnLayoutInput, tableFormatInput } from './parser-model.js';
import { resolveTableColumnWidths } from './layout/table-columns.js';
import {
  measureParagraphIntrinsicWidths,
  measureTableCellIntrinsicWidths,
} from './layout/intrinsic-width.js';


export { computeColumns };

// ── Line-layout engine (segmentation + line-breaking + measurement) ──────────
// Lifted into ./line-layout.ts (verbatim, B2 phase boundary). renderer.ts is the
// paint/paginate side; it drives the pure kernel below. One-directional import
// (renderer → line-layout); line-layout imports RenderState/DecodedImage back as
// a TYPE only (erased), so there is no runtime cycle.
import {
  DEFAULT_TAB_PT,
  buildFont,
  buildSegments,
  fontClassesWithPitches,
  getDefaultFontSize,
  gridCharDeltaPx,
  gridSegDeltaPx,
  hasCJKBreakOpportunity,
  isGridLineRule,
  layoutLines,
  lineBoxHeight,
  normalizeFontFamilyUncached,
  paragraphMarkLineHeight,
  paragraphSegsStateSensitive,
  rescaleLayoutLines,
  rubyAscentReservePx,
  segAdvanceWidth,
  segLetterSpacingPx,
  shapeRenderState,
  segmentCharacterGridDeltaPx,
  segmentEastAsiaFloorSingleLinePx,
  segmentIntendedSingleLinePx,
  splitTextForLayout,
} from './line-layout.js';
import type {
  DocGridCtx,
  LineBoundary,
  LayoutImageSeg,
  LayoutLine,
  LayoutMathSeg,
  LayoutSeg,
  LayoutTabSeg,
  LayoutTextSeg,
  WrapLayoutCtx,
} from './line-layout.js';
import {
  emphasisMarkCenters,
  emphasisMarkGeometry,
} from './emphasis-mark.js';
import {
  createFloatWrapOracle,
  measureParagraph,
  type MeasuredParagraph,
  type ParagraphMeasurementEnvironment,
} from './paragraph-measure.js';
import {
  paragraphFragmentAdvancePt,
  type FlowFragment,
} from './layout/flow-fragment.js';
import {
  acquireRetainedTable,
  type RetainedTableAcquisition,
  type RetainedTableAcquisitionDependencies,
} from './layout/table-acquisition.js';
import { combineAdjacentTableLayoutInputs } from './layout/adjacent-table-layout-input.js';
import { layoutTable as layoutRetainedTableInput } from './layout/table.js';
import {
  startTableFragmentCursor,
  takeTableFragment,
  type PageDependentTableBlockRequest,
  type TableFragmentLayout,
} from './layout/table-pagination.js';
import { paragraphGapAdjustment, paragraphGapPt } from './layout/paragraph-spacing.js';
import { imageResourceKey } from './layout/source-key.js';
import {
  hasVisibleParagraphBorder as hasAnyBorderEdge,
  paragraphsShareBorderBox as parasShareBorderBox,
  resolveParagraphBorderEdges,
} from './layout/paragraph-border-adjacency.js';
import {
  acquireParagraphResult,
  acquireRetainedFrameGroup,
  acquireShapeTextBoxLayout,
  bodyFrameGroupFor,
  bodyParagraphBorderEdgesFor,
  paragraphLayoutFromMeasurement,
  sliceParagraphLayout,
  type CompleteTextBoxStoryAcquirer,
} from './layout/paragraph.js';
import {
  ownedParagraphAnchorCollisions,
  inheritedParagraphAuthorityForReacquisition,
  TRANSIENT_TABLE_FINAL_FRAME_EXCLUSION_PREFIX,
} from './layout/paragraph-wrap-registry.js';
import { acquireRegisteredParagraph } from './layout/registered-paragraph-acquisition.js';
import {
  paragraphAnchorCollisions,
  paragraphWrapExclusions,
} from './layout/paragraph-float-authority.js';
import { paragraphAnchorReferenceFrames } from './paragraph-anchor-frame-adapter.js';
import {
  applyDrawingMLCollisionRegistryDelta,
  createDrawingMLCollisionRegistry,
  drawingMLCollisionRegistryDelta,
  validateDrawingMLCollisionRegistryDelta,
} from './layout/drawingml-collision-registry.js';
import { resolveAnchorFrame } from './layout/anchor-frame.js';
import {
  drawVerticalRun,
  drawTateChuYokoRun,
  drawUprightBox,
  physicalToLogicalAnchorBox,
  verticalTextLayerPlacement,
} from './vertical-text.js';

const HIGHLIGHT_COLORS: Record<string, string> = {
  yellow: '#FFFF00', cyan: '#00FFFF', green: '#00FF00', magenta: '#FF00FF',
  blue: '#0000FF', red: '#FF0000', darkBlue: '#000080', darkCyan: '#008080',
  darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000',
  darkYellow: '#808000', darkGray: '#808080', lightGray: '#C0C0C0',
  black: '#000000', white: '#FFFFFF',
};

function kashidaLevelOf(alignment: string | null | undefined): KashidaLevel | null {
  if (alignment === 'lowKashida') return 'low';
  if (alignment === 'mediumKashida') return 'medium';
  if (alignment === 'highKashida') return 'high';
  return null;
}

/**
 * ECMA-376 §17.18.44 true-kashida allocation shared by body and Word-textbox
 * lines. The delta form pins the original string to layout's measuredWidth;
 * inserted tatweels then grow it under the same font, kerning, w:w scale,
 * character spacing, and character-grid model used by layout and paint.
 */
function computeLineKashidaDistribution(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  segments: readonly LayoutSeg[],
  slackPx: number,
  level: KashidaLevel,
  scale: number,
  fontFamilyClasses: Record<string, string>,
  gridDeltaPx: number,
) {
  const distSegs = segments.map((seg) =>
    'text' in seg && (seg as LayoutTextSeg).fitTextRegionIndex === undefined
      ? { text: (seg as LayoutTextSeg).text }
      : {},
  );
  const originalModelAdvance = new Map<number, number>();
  const modeledAdvance = (si: number, text: string): number => {
    const s = segments[si] as LayoutTextSeg;
    ctx.font = buildFont(
      s.bold,
      s.italic,
      calcEffectiveFontPx(s, scale),
      s.fontFamily,
      fontFamilyClasses,
      s.fontRoute,
    );
    const prevKerning = ctx.fontKerning;
    const prevLetterSpacing = ctx.letterSpacing;
    if (s.kerning != null) {
      ctx.fontKerning = s.fontSize >= s.kerning ? 'normal' : 'none';
    }
    // Layout measures natural glyph advance and folds fixed pitch in itself.
    ctx.letterSpacing = '0px';
    const naturalWidth = ctx.measureText(text).width;
    ctx.letterSpacing = prevLetterSpacing;
    if (s.kerning != null) ctx.fontKerning = prevKerning;
    return segAdvanceWidth({ ...s, text }, naturalWidth, gridDeltaPx, scale);
  };
  const measureAdvance = (si: number, text: string): number => {
    const s = segments[si] as LayoutTextSeg;
    let originalAdvance = originalModelAdvance.get(si);
    if (originalAdvance === undefined) {
      originalAdvance = modeledAdvance(si, s.text);
      originalModelAdvance.set(si, originalAdvance);
    }
    if (text === s.text) return s.measuredWidth;
    return s.measuredWidth + modeledAdvance(si, text) - originalAdvance;
  };
  return computeKashidaDistribution(distSegs, slackPx, level, measureAdvance);
}

function collectMathRuns(
  body: BodyElement[],
  story: import('./layout/types.js').SourceRef['story'] = 'body',
  storyInstance = 'body',
): ReturnType<typeof bodyMathOccurrences> {
  return bodyMathOccurrences(body, story, storyInstance);
}

/** True if any currently representable document story contains OMML. The body
 * array form remains supported for existing callers. */
export function documentHasMath(input: BodyElement[] | DocxDocumentModel): boolean {
  return (Array.isArray(input) ? collectMathRuns(input) : documentMathOccurrences(input)).length > 0;
}

/** Build the one immutable resource snapshot shared by pagination and paint.
 * Browser handles remain on this runtime adapter and never enter layout data. */
export function createLayoutServices(
  doc: DocxDocumentModel,
  options: {
    readonly localMetrics?: Readonly<Record<string, ResolvedLocalFontMetric>>;
    readonly useGoogleFonts?: boolean;
    readonly mathResources?: readonly MathLayoutResource[];
    readonly mathDrawables?: ReadonlyMap<string, CanvasImageSource>;
    readonly measureContext?: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D | null;
    /** Successfully loaded/registered faces only; declarations are not inventory. */
    readonly embeddedFaces?: readonly FontFace[];
    readonly googleFaces?: readonly FontFace[];
  } = {},
): LayoutServices {
  doc = normalizeInternalDocumentModel(doc).document;
  const localMetrics = snapshotLocalMetrics(options.localMetrics);
  const fontFamilyCharsets = Object.freeze(Object.fromEntries(
    Object.entries(internalDocumentModel(doc).fontFamilyCharsets ?? {})
      .map(([family, charset]) => [family.trim().toLowerCase(), charset]),
  ));
  const displayFaceFamily = (family: string): string => family
    .trim()
    .replace(/^(['"])(.*)\1$/, '$2');
  const normalizedFaceFamily = (family: string): string => displayFaceFamily(family)
    .toLocaleLowerCase('en-US');
  const loadedFaceStyle = (face: FontFace): 'normal' | 'italic' | null => {
    const style = face.style.trim().toLocaleLowerCase('en-US');
    return style === 'normal' || style === 'italic' ? style : null;
  };
  const loadedFaceWeight = (face: FontFace): number | null => {
    const weight = face.weight.trim().toLocaleLowerCase('en-US');
    if (weight === 'normal') return 400;
    if (weight === 'bold') return 700;
    if (!/^\d+$/.test(weight)) return null;
    const numeric = Number(weight);
    return numeric >= 100 && numeric <= 900 ? numeric : null;
  };
  const loadedFaces = (faces: readonly FontFace[]): Array<{ family: string; displayFamily: string; weight: number; style: 'normal' | 'italic' }> =>
    faces.flatMap((face) => {
      if (face.status !== 'loaded') return [];
      const weight = loadedFaceWeight(face);
      const style = loadedFaceStyle(face);
      return weight == null || style == null ? [] : [{
        family: normalizedFaceFamily(face.family),
        displayFamily: displayFaceFamily(face.family),
        weight,
        style,
      }];
    });
  const successfulEmbedded = new Map(loadedFaces(options.embeddedFaces ?? []).map((loaded) => [
    `${loaded.family}:${loaded.weight}:${loaded.style}`, loaded,
  ]));
  const inventory: FontInventoryFace[] = (doc.embeddedFonts ?? []).flatMap((font) => {
    const weight = font.style === 'bold' || font.style === 'boldItalic' ? 700 : 400;
    const style = font.style === 'italic' || font.style === 'boldItalic' ? 'italic' as const : 'normal' as const;
    const loaded = successfulEmbedded.get(`${normalizedFaceFamily(font.fontName)}:${weight}:${style}`);
    if (!loaded) return [];
    return [{
      requestedFamily: font.fontName,
      resolvedFamily: loaded.displayFamily,
      source: 'embedded' as const,
      weight,
      style,
    }];
  });
  for (const [requestedFamily, metric] of Object.entries(localMetrics)) {
    inventory.push({
      requestedFamily: metric.requestedFamily ?? requestedFamily,
      resolvedFamily: metric.family,
      source: 'local',
      weight: metric.weight ?? 400,
      style: metric.style ?? 'normal',
    });
  }
  if (options.useGoogleFonts) {
    const successfulGoogle = loadedFaces(options.googleFaces ?? []);
    const seen = new Set<string>();
    for (const name of docxFontPreloadNames(doc)) {
      if (!name) continue;
      const key = name.toLocaleLowerCase('en-US');
      if (seen.has(key)) continue;
      seen.add(key);
      const entry = DOCX_GOOGLE_FONTS[key];
      const resolvedFamily = entry?.loadFamily ?? name;
      if (!entry) continue;
      for (const loaded of successfulGoogle.filter((face) => face.family === normalizedFaceFamily(resolvedFamily))) {
        inventory.push({
          requestedFamily: name,
          resolvedFamily: loaded.displayFamily,
          source: normalizedFaceFamily(resolvedFamily) === normalizedFaceFamily(name) ? 'google' : 'substitute',
          weight: loaded.weight,
          style: loaded.style,
        });
      }
    }
  }
  const ctx = options.measureContext ?? (typeof OffscreenCanvas !== 'undefined'
    ? new OffscreenCanvas(1, 1).getContext('2d')
    : typeof document !== 'undefined'
      ? document.createElement('canvas').getContext('2d')
      : null);
  const routedFontFamilies = [...new Set([
    ...Object.keys(doc.fontFamilyClasses ?? {}),
    ...Object.keys(doc.fontFamilyPitches ?? {}),
    ...docxRenderedFontFamilies(doc),
    ...(doc.majorFont ? [doc.majorFont] : []),
    ...(doc.minorFont ? [doc.minorFont] : []),
  ])];
  const textBase = createTextLayoutService({
    fonts: createFontResolver(inventory, {
      nativeFamilyLists: Object.fromEntries(
        routedFontFamilies.map((family) => [
          family,
          normalizeFontFamilyUncached(
            family,
            doc.fontFamilyClasses ?? {},
            doc.fontFamilyPitches ?? {},
          ),
        ]),
      ),
    }),
    localMetrics,
    eastAsiaFontCharsets: fontFamilyCharsets,
    genericFamilies: Object.fromEntries(
      routedFontFamilies.map((family) => [
        family,
        classifyDocxFontGeneric(family, doc.fontFamilyClasses, doc.fontFamilyPitches),
      ]),
    ),
    measurer: {
      fingerprint: ctx ? 'canvas-text-metrics-v1' : 'deterministic-text-metrics-v1',
      measure(request: Readonly<GlyphMeasureRequest>) {
        if (!ctx) return {
          advancePt: [...request.text].length * request.fontSizePt * 0.5,
          ascentPt: request.fontSizePt * 0.8,
          descentPt: request.fontSizePt * 0.2,
        };
        const previousFont = ctx.font;
        const previousLetterSpacing = ctx.letterSpacing;
        const previousKerning = ctx.fontKerning;
        try {
          ctx.font = canvasFontString(request.fontRoute, request.fontSizePt, request.weight, request.style);
          ctx.letterSpacing = `${request.letterSpacingPt}px`;
          if (request.kerning != null) ctx.fontKerning = request.kerning ? 'normal' : 'none';
          const metrics = ctx.measureText(request.text);
          const inkBounds = {
            xMinPt: Number.isFinite(metrics.actualBoundingBoxLeft)
              ? -metrics.actualBoundingBoxLeft : 0,
            xMaxPt: Number.isFinite(metrics.actualBoundingBoxRight)
              ? metrics.actualBoundingBoxRight : metrics.width,
            ascentPt: metrics.actualBoundingBoxAscent,
            descentPt: metrics.actualBoundingBoxDescent,
          };
          const hasFiniteInkBounds = Object.values(inkBounds).every(Number.isFinite);
          return {
            advancePt: metrics.width,
            ascentPt: metrics.fontBoundingBoxAscent ?? metrics.actualBoundingBoxAscent ?? 0,
            descentPt: metrics.fontBoundingBoxDescent ?? metrics.actualBoundingBoxDescent ?? 0,
            ...(hasFiniteInkBounds ? { inkBounds } : {}),
          };
        } finally {
          ctx.font = previousFont;
          ctx.letterSpacing = previousLetterSpacing;
          if (request.kerning != null) ctx.fontKerning = previousKerning;
        }
      },
    },
  });
  const mathOccurrences = normalizeInternalDocumentModel(doc).mathOccurrences;
  const mathResources = options.mathResources ?? mathOccurrences.map(({ display, source }) => ({
    resourceKey: mathResourceKey(source, display ? 'display' : 'inline'),
    widthEm: 0,
    ascentEm: 0,
    descentEm: 0,
    available: false,
    diagnostics: [{
      code: 'UNSUPPORTED_FEATURE' as const,
      severity: 'warning' as const,
      message: 'The optional DOM math engine is unavailable; using the worker-safe text fallback',
    }],
  }));
  const imageMetadata = documentImageMetadataRecords(doc, (paragraph) => {
    const numbering = paragraph.numbering;
    if (!numbering) throw new Error('Picture-bullet metadata requires numbering');
    const marker = numberingMarkerShapeInput(numbering, getDefaultFontSize(paragraph));
    return {
      widthPt: numbering.picBulletWidthPt ?? marker.fontSizePt,
      heightPt: numbering.picBulletHeightPt ?? marker.fontSizePt,
    };
  });
  const services: LayoutServices = Object.freeze({
    text: textBase,
    images: createImageMetadataService(imageMetadata),
    math: createMathMetadataService(mathResources),
  });
  const occurrenceKeys = mathOccurrences.map(({ source, display }) =>
    mathResourceKey(source, display ? 'display' : 'inline'));
  const metadataKeys = mathResources.map((resource) => resource.resourceKey);
  const missingMetadata = occurrenceKeys.filter((key) => !metadataKeys.includes(key));
  const extraMetadata = metadataKeys.filter((key) => !occurrenceKeys.includes(key));
  if (missingMetadata.length || extraMetadata.length) {
    throw new Error(
      `Math metadata membership mismatch: missing [${missingMetadata.join(', ')}]; extra [${extraMetadata.join(', ')}]`,
    );
  }
  const availableMathKeys = mathResources
    .filter((resource) => resource.available !== false)
    .map((resource) => resource.resourceKey);
  attachPrivateResourceLookup(
    services,
    options.mathDrawables ?? new Map(),
    availableMathKeys,
  );
  attachPaintResourceRegistry(services, createDocumentPaintResourceRegistry(doc, imageMetadata));
  attachBodyLayoutKernel(services, createConcreteBodyLayoutKernel(doc, ctx, localMetrics));
  return services;
}

/** Rasterize an SVG string to an <img> (browser). Resolves once decoded. */
function svgToImage(svg: string): Promise<HTMLImageElement> {
  const url = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  const img = new Image();
  return new Promise((resolve, reject) => {
    img.onload = () => resolve(img);
    img.onerror = reject;
    img.src = url;
  });
}

/** Convert equations before layout. Math resources use only normalized,
 * structural SourceRef/resourceKey facts; parser object identity is irrelevant. */
export async function prepareMathRuns(
  input: BodyElement[] | DocxDocumentModel,
  math: MathRenderer,
) {
  if (Array.isArray(input)) {
    throw new TypeError('prepareMathRuns requires a document model so every story has an explicit structural source');
  }
  const runs = normalizeInternalDocumentModel(input).mathOccurrences;
  if (runs.length === 0) return { records: [], drawables: new Map() };
  await math.loadMathJax();
  const records: MathLayoutResource[] = [];
  const drawables = new Map<string, CanvasImageSource>();
  const seen = new Set<string>();
  for (const r of runs) {
    const resourceKey = r.resourceKey;
    if (seen.has(resourceKey)) throw new Error(`Duplicate math occurrence: ${resourceKey}`);
    seen.add(resourceKey);
    try {
      const out = await math.mathMLToSvg(mathToMathML(r.nodes, r.display));
      const img = await svgToImage(recolorSvg(out.svg, '#000000'));
      records.push({
        resourceKey,
        widthEm: out.widthEm,
        ascentEm: out.ascentEm,
        descentEm: out.descentEm,
        diagnostics: [],
      });
      drawables.set(resourceKey, img);
    } catch {
      records.push({
        resourceKey,
        widthEm: 0,
        ascentEm: 0,
        descentEm: 0,
        available: false,
        diagnostics: [{
          code: 'UNSUPPORTED_FEATURE',
          severity: 'warning',
          message: 'Math conversion failed; using the deterministic text fallback',
        }],
      });
    }
  }
  return { records, drawables };
}

interface RetainedTableRecord {
  readonly sourceIndex: number;
  readonly acquisition: RetainedTableAcquisition;
  readonly contentWidthPt: number;
  readonly anchorYPt: number;
}

export interface RenderState {
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
  scale: number;    // px per pt
  /** Device-pixel ratio the canvas was scaled by (`ctx.scale(dpr, dpr)`). Used to
   *  compute the crisp-line offset (see crispOffset) so thin axis-aligned strokes
   *  land on a single device row instead of straddling two. */
  dpr: number;
  /** Retained point-space to final logical CSS, including the active page frame. */
  pointToCss?: Matrix2DData;
  contentX: number; // left of content area (px)
  contentW: number; // width of content area (px)
  y: number;        // current Y cursor (px)
  pageH: number;    // full page height (px)
  defaultColor: string;
  /** 0-based page index currently being rendered */
  pageIndex: number;
  /** total page count in the document */
  totalPages: number;
  /** ECMA-376 §17.6.12 — the DISPLAYED page number for the current page (after
   *  per-section `w:start` restart). A PAGE field renders this instead of the raw
   *  `pageIndex + 1`. Absent ⇒ `pageIndex + 1` (single-section fallback). */
  displayPageNumber?: number;
  /** ECMA-376 §17.6.12 / §17.18.59 — the ST_NumberFormat governing the current
   *  page's number (from the page's section `w:fmt`). A PAGE field formats its
   *  result with this unless it carries its own `\*` switch (§17.16.4.3.1). Absent
   *  ⇒ decimal. */
  pageNumberFormat?: NumberFormat;
  /** preloaded drawable images keyed by `imageKey(imagePath, colorReplaceFrom)`
   *  (raster ImageBitmap or, for an `asvg:svgBlip` vector original, an
   *  HTMLImageElement) */
  images: Map<string, DecodedImage>;
  /** when true, layout is performed but nothing is drawn (used for header/footer height measurement) */
  dryRun: boolean;
  /** section left margin in pt — used to convert margin-relative anchor X to page-absolute */
  marginLeft: number;
  /** section right margin in pt — used by anchor positioning to resolve
   *  `<wp:positionH relativeFrom="margin">` and the `*Margin` family containers. */
  marginRight: number;
  /** ECMA-376 §17.6.11: the body's TOP/BOTTOM **inset** from the page edge in pt — the
   *  margin's MAGNITUDE (|margin|), NOT the signed pgMar value. A negative top/bottom
   *  margin (ST_SignedTwipsMeasure) measures the body |margin| from the page edge and
   *  overlaps the header/footer; `bodyMarginInsetPt` derives this at the writers
   *  (baseState, buildMeasureState). Read as the column region top (renderBodyElements /
   *  splitParagraphAcrossPages) and as the text-margin container for `relativeFrom=
   *  "topMargin"/"bottomMargin"/"margin"` anchors/frames (anchor-geometry, frame-geometry;
   *  §17.18.100 — the text-margin location IS the body edge). Do NOT treat as the signed
   *  margin: the overflow decision keeps the sign separately (header/footerOverflowPt). */
  marginTop: number;
  marginBottom: number;
  /** Section page width in pt. */
  pageWidth: number;
  /** Active anchor-image floats that constrain text layout on the current page. */
  floats: FloatRect[];
  /** Monotonic counter assigning a unique id to each registerAnchorFloats call,
   *  i.e. one id per paragraph per page. Used only to scope the implementation-
   *  defined (HEURISTIC) overlap avoidance to DIFFERENT paragraphs. Reset to 0
   *  on every page flip so measure and render assign matching paraIds. */
  floatParaSeq: number;
  /** ECMA-376 §17.6.5 docGrid (type + pitch), applied to auto line spacing. */
  docGrid: DocGridCtx;
  /** Document-wide OOXML layout policy normalized once at renderer entry. */
  layoutSettings: DocumentLayoutSettings;
  /** Active section geometry and grid policy normalized for this state. */
  sectionLayout: SectionLayoutContext;
  /** Active WordprocessingML story and nested text-container stack. */
  storyContext: StoryContext;
  /** True when the document body contains East Asian text. Gates docGrid line-
   *  cell rounding of empty / anchor-only paragraph marks (see
   *  paragraphMarkLineHeight), which carry no text to classify themselves. */
  docEastAsian: boolean;
  /** ECMA-376 §17.8.3.10 — font→family map from word/fontTable.xml. Used by
   *  resolveFontFamily as the authoritative source of serif/sans-serif classification. */
  fontFamilyClasses: Record<string, string>;
  /** Exact local faces and version-adaptive Word line metrics resolved before
   * pagination. Keyed by normalized authored family. */
  resolvedLocalFonts: Readonly<Record<string, ResolvedLocalFontMetric>>;
  /** Instance-scoped resource snapshot shared by pagination and paint. */
  layoutServices?: LayoutServices;
  /** Per-render opaque retained resource session adapter. */
  retainedResourcePainter?: CanvasPaintResourcePainter;
  /** Renderer authorities injected into the pure recursive table acquisition. */
  retainedTableAcquisition?: RetainedTableAcquisitionDependencies<RenderState>;
  /** Session-local bridge from shape geometry to the shared paragraph/table
   * story algorithms. Kept renderer-private so retained layouts stay plain. */
  acquireCompleteTextBoxStory?: CompleteTextBoxStoryAcquirer;
  /** Each body occurrence owns its acquisition and pagination context for this
   * render session. Parser objects may legally occur more than once. */
  retainedTablesBySourceIndex?: Map<number, RetainedTableRecord>;
  /** ECMA-376 §17.15.1.58–.60 — resolved Japanese line-breaking rules
   *  (kinsoku enabled flag + line-start/line-end forbidden character sets).
   *  Default is the application's Japanese kinsoku table with kinsoku ON. */
  kinsoku: KinsokuRules;
  /** ECMA-376 §17.15.1.25 `w:defaultTabStop` — the interval (points) at which
   *  automatic tab stops are generated after all custom stops. Threaded from
   *  `doc.settings.defaultTabStop` like `kinsoku` so the MEASURE pass matches
   *  the DRAW pass; falls back to {@link DEFAULT_TAB_PT} (720 twips = 36pt) when
   *  the document omits the element. */
  defaultTabPt: number;
  /** ECMA-376 §17.15.1.18 — East Asian punctuation / character-spacing mode. */
  characterSpacingControl?: string;
  /** ECMA-376 §17.15.3.1 `w:compat/w:useFELayout`. */
  useFeLayout?: boolean;
  /** ECMA-376 §17.15.3.1 `w:compat/w:balanceSingleByteDoubleByteWidth`. */
  balanceSingleByteDoubleByteWidth?: boolean;
  /** ECMA-376 §22.1.2.30 `m:mathPr/m:defJc` — document-wide default math
   *  justification (ST_Jc math). `undefined` ⇒ spec default `centerGroup`.
   *  Threaded from `doc.settings.mathDefJc` like `kinsoku`; consumed by the
   *  per-line alignment step for single display-math lines. */
  mathDefJc?: string;
  /** Callback for building a transparent text selection overlay. */
  onTextRun?: (run: DocxTextRunInfo) => void;
  /** When false, runs tagged with a `revision` render without the
   *  track-changes overlay (no author colour, no underline/strikethrough). */
  showTrackChanges: boolean;
  /** ECMA-376 §17.16.5.16 DATE / §17.16.5.72 TIME — the "current" instant (epoch
   *  ms) a DATE/TIME field formats through its `\@` picture (§17.16.4.1). Injected
   *  from the render option `currentDate` so field output is deterministic under
   *  test; absent ⇒ `Date.now()` (real time). */
  currentDateMs?: number;
  /** ECMA-376 §17.11 — footnote/endnote reference markers (`noteRef` runs)
   *  display the note's 1-based sequential number, not the raw `@w:id`. Keyed by
   *  `"footnote:<id>"` / `"endnote:<id>"`. The in-note `<w:footnoteRef>`
   *  placeholder (empty id) is substituted with the number provided via
   *  {@link noteReferenceNumber} while laying out that note's content. */
  noteNumbers?: Map<string, number>;
  /** Set while laying out a footnote/endnote's own content, so the leading
   *  `<w:footnoteRef>` placeholder (which carries no id) renders the note's
   *  number. Undefined for body text. */
  noteReferenceNumber?: number;
  /** ECMA-376 §20.4.3.2/§20.4.3.5: a DrawingML anchor whose `<wp:positionV>`
   *  uses a page-level `relativeFrom` (page / margin / topMargin / bottomMargin
   *  / leftMargin / rightMargin / insideMargin / outsideMargin / column) is
   *  positioned independently of its source-order anchoring paragraph — Word
   *  lays it out as soon as the page is opened, so paragraphs that come BEFORE
   *  the anchor's paragraph in source order still wrap around it. To match,
   *  we pre-scan upcoming body paragraphs at every page-start and register
   *  these floats up front. This set records which paragraphs have had their
   *  page-level floats pre-registered on the current page, so
   *  {@link registerAnchorFloats} skips re-registering them when the main
   *  flow reaches that paragraph. Reset whenever floats are reset (page flip
   *  or column relocation that rolls back this paragraph's own floats). */
  pageAnchorPrescanned?: Set<DocParagraph>;
  /** ECMA-376 §17.6.20 vertical writing — the FRAME-level vertical flag. When
   *  true the page is laid out in a SWAPPED logical coordinate space (logical
   *  width = physical page height) and the canonical section-region matrix rotates
   *  body paint +90° into physical space. On a tbRl-family
   *  page (no {@link verticalAllRotated}) the glyph-draw path then counter-
   *  rotates each upright (CJK) glyph −90° about its own centre so ideographs
   *  stand upright while Latin/digits stay sideways (correct for vertical
   *  Japanese); a `btLr` page sets {@link verticalAllRotated} too, which
   *  SUPPRESSES that counter-rotation (all glyphs ride the region transform).
   *  Absent / false ⇒ horizontal (lrTb) — the whole layout + paint path is
   *  byte-identical to the pre-vertical renderer. */
  verticalCJK?: boolean;
  /** ECMA-376 §17.6.20 + Part 4 §14.11.7 — set (always together with
   *  {@link verticalCJK}) on a section-level `btLr` page. Word GT (issue #988
   *  re-adjudication, raster-proven on asymmetric glyphs — the dakuten of 「び」
   *  lands bottom-right): `btLr` shares the `tbRl` PAGE FRAME (swapped logical
   *  layout, +90° page paint, columns right→left) but rotates EVERY glyph with
   *  the page — CJK is NOT counter-rotated upright, vertical punctuation forms
   *  (、。（） → U+FE1x/FE3x) are NOT substituted, and 縦中横 (§17.3.2.10) is
   *  NOT grouped. When set, the glyph-draw sites take the ordinary HORIZONTAL
   *  branches inside the rotated frame, so the page raster equals the
   *  horizontal rendering of the swapped frame rotated +90° CW wholesale.
   *  Frame-level consumers (region transform, text-layer transform, `verticalPhys`
   *  anchors, upright images/tables — GT-less for btLr, kept at tbRl parity)
   *  keep reading {@link verticalCJK}. Absent ⇒ upright-vertical (tbRl family)
   *  or horizontal. */
  verticalAllRotated?: boolean;
  /** ECMA-376 §17.6.20 + §20.4.3.x — the PHYSICAL page geometry for a vertical
   *  (tbRl) page, in the SAME units the rest of RenderState uses (margins/page
   *  size in pt; `cssWidthPx` in px). Present only when `verticalCJK` is set.
   *  A DrawingML anchor's `<wp:positionH/V>` is resolved against the PHYSICAL page
   *  (the drawing layer is placed independently of the text-flow rotation), so the
   *  anchor path builds a PHYSICAL-geometry proxy RenderState from this and maps
   *  the resolved physical box into the swapped logical frame via
   *  {@link physicalToLogicalAnchorBox}. The four margins are the physical
   *  pgMar values (already the body inset for the top/bottom, matching
   *  `bodyMarginInsetPt`). `cssWidthPx` = physical page width in px = the page
   *  transform's `translate(cssWidth, 0)` term. Absent ⇒ horizontal. */
  verticalPhys?: {
    pageWidth: number;
    pageHeight: number;
    marginLeft: number;
    marginRight: number;
    marginTop: number;
    marginBottom: number;
    cssWidthPx: number;
  };
  /** ECMA-376 §17.3.2.6 — the effective background (hex 6, no `#`) behind the text
   *  from the ENCLOSING containers, most-specific first: a table cell's `<w:tcPr>
   *  <w:shd w:fill>` (§17.4.33), overridden by a paragraph's `<w:pPr><w:shd w:fill>`
   *  (§17.3.1.31). An automatic run color (`<w:color w:val="auto"/>`, no explicit
   *  color) contrasts against this when the run has no closer background of its own
   *  (its run-level `<w:shd>`). Threaded by `renderCell` (from the cell fill) and
   *  `renderParagraph` (paragraph shading overrides); absent ⇒ the page background.
   *  Only the auto-contrast decision reads it — it does NOT paint any rect (cell /
   *  paragraph shading rects are painted by their own passes). */
  containerShading?: string | null;
}

const BODY_STORY_CONTEXT: StoryContext = {
  story: 'body',
  containers: [],
  lineNumberingEligible: true,
};

export function resolveBodyParagraphLayoutContext(
  state: Pick<RenderState, 'layoutSettings' | 'sectionLayout'>
    & Partial<Pick<RenderState, 'layoutServices' | 'defaultTabPt'>>,
  paragraph: DocParagraph,
): ParagraphLayoutContext {
  const context = resolveParagraphLayoutContext(
    state.layoutSettings,
    state.sectionLayout,
    BODY_STORY_CONTEXT,
    paragraph,
  );
  return applyNumberingBodyOffset(context, {
    numbering: paragraph.numbering,
    ...(paragraph.numbering ? {
      markerInput: numberingMarkerShapeInput(paragraph.numbering, getDefaultFontSize(paragraph)),
    } : {}),
    authoredFirstIndentPt: paragraph.indentFirst,
    tabStops: paragraph.tabStops,
    defaultTabPt: state.defaultTabPt,
    service: state.layoutServices?.text,
  });
}

function resolveStateParagraphLayoutContext(
  state: Pick<RenderState, 'layoutSettings' | 'sectionLayout' | 'storyContext'>
    & Partial<Pick<RenderState, 'layoutServices' | 'defaultTabPt'>>,
  paragraph: DocParagraph,
): ParagraphLayoutContext {
  const context = resolveParagraphLayoutContext(
    state.layoutSettings,
    state.sectionLayout,
    state.storyContext ?? BODY_STORY_CONTEXT,
    paragraph,
  );
  return applyNumberingBodyOffset(context, {
    numbering: paragraph.numbering,
    ...(paragraph.numbering ? {
      markerInput: numberingMarkerShapeInput(paragraph.numbering, getDefaultFontSize(paragraph)),
    } : {}),
    authoredFirstIndentPt: paragraph.indentFirst,
    tabStops: paragraph.tabStops,
    defaultTabPt: state.defaultTabPt,
    service: state.layoutServices?.text,
  });
}

function withTableCellStory(state: RenderState): RenderState {
  return {
    ...state,
    storyContext: enterTableCellStoryContext(
      state.storyContext ?? BODY_STORY_CONTEXT,
    ),
  };
}

/** Whether a paragraph may use scale-1 glyph geometry as the single layout
 * authority and map it to the paint viewport with a Canvas transform.
 *
 * Keep this predicate shared by paint and table-cell height measurement: row
 * height / vAlign must reserve the exact line boxes that the glyph path draws.
 * The excluded paths still have scale-aware layout or decoration behavior —
 * including marker/tab-leader paint and math fallback — that has not yet been
 * expressed entirely in canonical document coordinates. */
function canonicalParagraphTextScaleEligible(
  storyContext: StoryContext,
  verticalCJK: boolean | undefined,
  inFrame: boolean,
  hasWrapContext: boolean,
  paragraphContext: Pick<ParagraphLayoutContext, 'hasRuby' | 'baseRtl'>,
  paragraph: Pick<DocParagraph, 'alignment' | 'numbering'>,
  segments: readonly LayoutSeg[],
): boolean {
  // `containers=[]` is the deliberate top-level body case. Nested body text is
  // accepted only while every enclosing story container is a table cell; other
  // stories/containers keep their established paint-space paths.
  const isSupportedBodyContainerChain =
    storyContext.story === 'body'
    && (storyContext.containers.length === 0
      || storyContext.containers.every((container) => container.kind === 'tableCell'));
  return !hasWrapContext
    && !inFrame
    && isSupportedBodyContainerChain
    && !verticalCJK
    && !paragraphContext.hasRuby
    && !paragraphContext.baseRtl
    && paragraph.numbering == null
    && !segmentsHaveRtl(segments)
    && kashidaLevelOf(paragraph.alignment) === null
    && segments.every((segment) =>
      !('isTab' in segment)
      && !('mathNodes' in segment)
      && (!('text' in segment) || segment.emphasisMark == null));
}

/** Information about a rendered text segment for building a transparent selection overlay. */
export interface DocxTextRunInfo {
  text: string;
  /** Left edge in canvas CSS px. */
  x: number;
  /** Top of line box in canvas CSS px. */
  y: number;
  /** Measured text width in CSS px. */
  w: number;
  /** Line height in CSS px. */
  h: number;
  /** Font size in CSS px. */
  fontSize: number;
  /** CSS `font` shorthand used for canvas drawing (e.g. `"bold 16px Arial"`). */
  font: string;
  /** Uniform per-code-point pitch in CSS px used to draw a horizontal run.
   *  Absent when the pitch is zero or the run uses vertical / 縦中横 paint. */
  letterSpacingPx?: number;
  /** ECMA-376 §17.6.20 (tbRl) — when the page is vertical the canvas is the
   *  physical landscape page rotated +90° at paint, so this run's `x`/`y` are the
   *  PHYSICAL top-left the overlay span must sit at, and `transform` is the CSS
   *  rotation (`"rotate(90deg)"`, applied about the span's top-left) that lays the
   *  horizontal DOM span along the drawn (rotated) glyph run. Absent for
   *  horizontal pages (the span is placed at `x`/`y` untransformed). */
  transform?: string;
  /** IX1 — the resolved hyperlink target of this run (ECMA-376 §17.16.22
   *  external URL / §17.16.23 internal `w:anchor` bookmark), or absent for a
   *  non-link run. The text-layer overlay turns a run carrying this into a
   *  clickable region; the drawn glyphs are unaffected. */
  hyperlink?: HyperlinkTarget;
  /** ECMA-376 §17.3.2.10 eastAsianLayout `w:vert` (縦中横 / horizontal-in-vertical):
   *  `true` when this run was drawn as tate-chu-yoko — its glyphs laid out
   *  horizontally, side by side, COMPRESSED into ONE em cell of the vertical
   *  column (see {@link drawTateChuYokoRun}). `w` is the drawn cell extent (one
   *  em), NOT the natural text width, so the find / selection overlays must clamp
   *  their horizontal extent to `w` rather than re-measuring the run's natural
   *  glyphs (issue #836). Absent for every ordinary run. */
  eastAsianVert?: boolean;
}

export interface RenderDocumentOptions {
  width?: number;
  dpr?: number;
  defaultTextColor?: string;
  /**
   * Lazy image-byte loader: fetch the raw bytes for an embedded image by zip
   * path, wrapped in a Blob of the given MIME (twin of pptx's `fetchImage`).
   * Supplied by {@link DocxDocument} (routing to its `getImage`), so the
   * renderer decodes images on demand instead of from inlined base64. When
   * omitted, images are skipped (no byte source).
   */
  fetchImage?: (path: string, mimeType: string) => Promise<Blob>;
  /** Called for each rendered text segment. Used to build a transparent text selection overlay. */
  onTextRun?: (run: DocxTextRunInfo) => void;
  /** Default `true`. When false, runs tagged with a `revision` (insertion or
   *  deletion from `<w:ins>` / `<w:del>`) render in their normal colour with
   *  no underline / strikethrough overlay — useful for a "final / no markup"
   *  view of a tracked document. */
  showTrackChanges?: boolean;
  /** ECMA-376 §17.16.5.16 DATE / §17.16.5.72 TIME — the "current" instant that a
   *  DATE/TIME field formats through its `\@` date picture (§17.16.4.1). Accepts a
   *  `Date` or epoch-ms number. Default = the real current time (`Date.now()` at
   *  render). Provide a fixed value to make DATE/TIME field output deterministic
   *  (e.g. in tests / reproducible exports). */
  currentDate?: Date | number;
  /** Internal per-document service snapshot. Public render options never expose it. */
  layoutServices?: LayoutServices;
  /** Internal load-time default captured once and mirrored into worker mode. */
  defaultCurrentDateMs?: number;
}

// ===== Image preloading =====

/**
 * A decoded, drawable image. Raster blips decode to an `ImageBitmap`
 * (createImageBitmap); the Microsoft `asvg:svgBlip` vector original decodes to
 * an `HTMLImageElement` (via core's path-keyed `getCachedSvgImageByPath`, since
 * `createImageBitmap` cannot rasterize SVG in every browser). Both are valid
 * `ctx.drawImage` sources with numeric `.width`/`.height`, so every draw site is
 * identical regardless of which kind was decoded.
 */
export type DecodedImage = ImageBitmap | HTMLImageElement;

interface ImagePair {
  /** Zip path of the raster fallback (or the SVG part itself when no raster
   *  blip is embedded). The cache key + the byte-fetch path. */
  imagePath: string;
  /** MIME type of the blip at {@link ImagePair.imagePath}. */
  mimeType: string;
  /**
   * Zip path of the vector original from the `asvg:svgBlip` extension, when
   * present. Preferred over `imagePath`; the decoded image is still stored
   * under `imagePath`'s key so draw sites (which look up by `imagePath`) find
   * it unchanged.
   */
  svgImagePath?: string;
  colorReplaceFrom?: string;
  /** ECMA-376 §20.1.8.23 `<a:duotone>` recolour, resolved to its two endpoint
   *  colours. When set, the decode remaps the raster along the `clr1`→`clr2`
   *  luminance ramp; the map key includes both colours so a duotone picture is
   *  cached separately from the raw blip. */
  duotone?: Duotone;
  /**
   * Largest intended draw size (pt) over every reference to this key. Only used
   * to pick a raster target resolution for vector metafiles (WMF/EMF), which
   * have no intrinsic pixel size — the player must rasterize at a chosen size.
   * Raster (PNG/JPEG) and SVG paths ignore it (they carry/scale their own
   * resolution). Defaults to 0 when no size is known.
   */
  widthPt: number;
  heightPt: number;
  /** True when at least one reference to this image carries an `<a:srcRect>`
   *  crop, so the decode must prefer the raster (the crop math needs the
   *  bitmap's native pixel grid; an SVG vector original has none). */
  hasCrop?: boolean;
}

/** Returns a stable map key for an (imagePath, colorReplaceFrom, duotone)
 *  triple. A plain picture is keyed by its zip path; an `a:clrChange`
 *  (colorReplaceFrom) and/or a `<a:duotone>` each append a suffix, so a
 *  recoloured variant is cached and looked up separately from the raw blip and
 *  from any other recolour combination. */
function imageKey(imagePath: string, colorReplaceFrom?: string, duotone?: Duotone): string {
  let key = imagePath;
  if (colorReplaceFrom) key += `|clr:${colorReplaceFrom}`;
  if (duotone) key += `|duo:${duotone.clr1}:${duotone.clr2}`;
  return key;
}

type DocxFetchImage = (path: string, mime: string) => Promise<Blob>;

// Second-layer cache for a picture's RECOLOUR result — the `a:clrChange`
// (colorReplaceFrom, §20.1.8.11) make-transparent pass and/or the `<a:duotone>`
// (§20.1.8.23) luminance ramp. The core path-keyed cache (getCachedBitmapByPath)
// holds the recolour-FREE bitmap — shared across every reference to a path and
// reclaimed with the document. The recolour pass (getImageData + putImageData,
// expensive) then runs once per (imagePath, colorReplaceFrom, duotone) triple
// and its ImageBitmap is kept here, so revisiting a page re-runs neither the
// decode NOR the recolour.
//
// Keyed FIRST by the document's `fetchImage` closure (one stable identity per
// DocxDocument), then by imageKey(imagePath, colorReplaceFrom, duotone) —
// mirroring the core cache's per-document namespacing so two documents sharing a
// zip path + recolour don't cross-contaminate, and the whole map is reclaimed
// with the document. The stored value is an ImageBitmap (a fresh OffscreenCanvas
// raster), so on destroy it must be closed (see dropColorReplacedCache), the
// same GPU-lifecycle discipline the core cache follows through its promise.
const colorReplacedByFetch = new WeakMap<DocxFetchImage, Map<string, Promise<ImageBitmap>>>();

function colorReplacedCacheFor(fetchImage: DocxFetchImage): Map<string, Promise<ImageBitmap>> {
  let cache = colorReplacedByFetch.get(fetchImage);
  if (!cache) {
    cache = new Map();
    colorReplacedByFetch.set(fetchImage, cache);
  }
  return cache;
}

/**
 * Close every color-replaced ImageBitmap for one document's `fetchImage` and
 * forget the document. Call from `DocxDocument.destroy()` alongside
 * `dropBitmapCacheByPath` (base bitmaps) and `dropSvgImageCache` (SVG object
 * URLs) so all three per-document image caches release promptly. A no-op when no
 * clrChange image was decoded. While a render pass holds a lease on this
 * document (core `acquireBitmapCacheLease`), the closes are deferred to the last
 * release — the same contract as the shared base/duotone caches — so a drop
 * racing an in-flight render never closes a bitmap mid-draw.
 */
export function dropColorReplacedCache(fetchImage: DocxFetchImage): void {
  const cache = colorReplacedByFetch.get(fetchImage);
  if (!cache) return;
  for (const p of cache.values()) deferBitmapCloseWhileLeased(fetchImage, p);
  cache.clear();
  colorReplacedByFetch.delete(fetchImage);
}

/** Picks a stable colour for a track-changes author. Mirrors Word's behaviour
 *  of cycling through a fixed palette (Word uses 8 hues then alternates).
 *  An empty / missing author maps to the first colour. */
const TRACK_CHANGE_AUTHOR_PALETTE = [
  '#C00000', // red
  '#0070C0', // blue
  '#00B050', // green
  '#7030A0', // purple
  '#E97132', // orange
  '#196B24', // dark green
  '#9E480E', // brown
  '#525252', // grey
];
function authorColor(author?: string): string {
  if (!author) return TRACK_CHANGE_AUTHOR_PALETTE[0];
  // Simple FNV-1a style hash so the same author always gets the same colour.
  let h = 0x811c9dc5;
  for (let i = 0; i < author.length; i++) {
    h ^= author.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return TRACK_CHANGE_AUTHOR_PALETTE[Math.abs(h) % TRACK_CHANGE_AUTHOR_PALETTE.length];
}

function collectImagePairs(doc: DocxDocumentModel, layoutServices: LayoutServices): ImagePair[] {
  const seen = new Map<string, ImagePair>();
  // Record one image reference (collapsing duplicate keys, tracking the max
  // intended draw size so a vector metafile is rasterized sharply enough for its
  // largest occurrence — only meaningful for WMF/EMF).
  const record = (pair: ImagePair) => {
    const key = imageKey(pair.imagePath, pair.colorReplaceFrom, pair.duotone);
    const existing = seen.get(key);
    if (!existing) {
      seen.set(key, pair);
    } else {
      existing.widthPt = Math.max(existing.widthPt, pair.widthPt);
      existing.heightPt = Math.max(existing.heightPt, pair.heightPt);
      // If ANY reference is cropped, force the raster decode for this key.
      existing.hasCrop = existing.hasCrop || pair.hasCrop;
    }
  };
  // ECMA-376 §17.9.9/§17.9.20 — a level's picture-bullet marker is an image
  // that lives on the paragraph's numbering, not in any run. Feed it into the
  // same decode pipeline (keyed by its zip path) so the marker draw site finds
  // a decoded bitmap.
  const recordPara = (para: DocParagraph, source: SourceRef) => {
    const num = para.numbering;
    const pb = num?.picBulletImagePath;
    if (pb && num) {
      const size = layoutServices.images.resolve(imageResourceKey(source, pb));
      record({
        imagePath: pb,
        mimeType: num.picBulletMimeType ?? '',
        widthPt: size.widthPt,
        heightPt: size.heightPt,
      });
    }
  };
  const walk = (runs: DocRun[]) => {
    for (const run of runs) {
      if (run.type === 'image') {
        const img = run as unknown as ImageRun;
        record({
          imagePath: img.imagePath,
          mimeType: img.mimeType,
          svgImagePath: img.svgImagePath,
          colorReplaceFrom: img.colorReplaceFrom,
          duotone: img.duotone,
          ...metafileRasterSize(img.mimeType, img.srcRect, img.widthPt ?? 0, img.heightPt ?? 0),
          hasCrop: img.srcRect != null,
        });
      } else if (run.type === 'shape') {
        // Inline images living inside a text box (<wps:txbx>) ride on the
        // shape's text blocks. Feed them into the same decode pipeline so the
        // WMF/EMF/raster/SVG decoders see their bytes (no colorReplace here).
        const shp = run as unknown as ShapeRun;
        for (const block of shp.textBlocks ?? []) {
          if (block.imagePath) {
            record({
              imagePath: block.imagePath,
              mimeType: block.mimeType ?? '',
              svgImagePath: block.svgImagePath,
              widthPt: block.imageWidthPt ?? 0,
              heightPt: block.imageHeightPt ?? 0,
            });
          }
        }
      }
    }
  };
  const walkTable = (
    tbl: DocTable,
    story: SourceRef['story'],
    storyInstance: string,
    prefix: readonly number[],
  ) => {
    for (let rowIndex = 0; rowIndex < tbl.rows.length; rowIndex += 1) {
      const row = tbl.rows[rowIndex]!;
      for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex += 1) {
        const cell = row.cells[cellIndex]!;
        for (let elementIndex = 0; elementIndex < cell.content.length; elementIndex += 1) {
          const ce = cell.content[elementIndex]!;
          const path = [...prefix, rowIndex, cellIndex, elementIndex];
          if (ce.type === 'paragraph') {
            const p = ce as unknown as DocParagraph;
            recordPara(p, { story, storyInstance, path });
            walk(p.runs);
          } else if (ce.type === 'table') {
            walkTable(ce as unknown as DocTable, story, storyInstance, path);
          }
        }
      }
    }
  };
  const walkBody = (
    body: BodyElement[],
    story: SourceRef['story'],
    storyInstance: string,
  ) => {
    for (let elementIndex = 0; elementIndex < body.length; elementIndex += 1) {
      const el = body[elementIndex]!;
      const source = { story, storyInstance, path: [elementIndex] } as const;
      if (el.type === 'paragraph') {
        const p = el as unknown as DocParagraph;
        recordPara(p, source);
        walk(p.runs);
      }
      if (el.type === 'table') walkTable(el as unknown as DocTable, story, storyInstance, source.path);
    }
  };
  walkBody(doc.body, 'body', 'body');
  if (doc.headers.default) walkBody(doc.headers.default.body, 'header', 'default');
  if (doc.headers.first)   walkBody(doc.headers.first.body, 'header', 'first');
  if (doc.headers.even)    walkBody(doc.headers.even.body, 'header', 'even');
  if (doc.footers.default) walkBody(doc.footers.default.body, 'footer', 'default');
  if (doc.footers.first)   walkBody(doc.footers.first.body, 'footer', 'first');
  if (doc.footers.even)    walkBody(doc.footers.even.body, 'footer', 'even');
  return [...seen.values()];
}

/**
 * Apply a:clrChange color replacement: turn every pixel whose (R,G,B) matches colorHex into
 * fully transparent (alpha=0). Returns a new ImageBitmap with the modified pixels.
 */
async function applyColorReplacement(bmp: ImageBitmap, colorHex: string): Promise<ImageBitmap> {
  const r = parseInt(colorHex.slice(0, 2), 16);
  const g = parseInt(colorHex.slice(2, 4), 16);
  const b = parseInt(colorHex.slice(4, 6), 16);

  const offscreen = new OffscreenCanvas(bmp.width, bmp.height);
  const ctx2 = offscreen.getContext('2d')!;
  ctx2.drawImage(bmp, 0, 0);

  const imgData = ctx2.getImageData(0, 0, bmp.width, bmp.height);
  const d = imgData.data;

  for (let i = 0; i < d.length; i += 4) {
    if (d[i] === r && d[i + 1] === g && d[i + 2] === b) {
      d[i + 3] = 0; // make transparent
    }
  }

  ctx2.putImageData(imgData, 0, 0);
  return createImageBitmap(offscreen);
}

/**
 * Decode a raster blip to an `ImageBitmap`, pulling the bytes lazily by zip path
 * via `fetchImage(imagePath, mimeType)` (twin of pptx's `fetchImage`) rather
 * than `fetch`-ing an inlined data URL. Applies an `a:clrChange`
 * (`colorReplaceFrom`) make-transparent pass when requested — unchanged
 * post-decode behavior.
 *
 * Two-layer caching so a page revisit re-runs NEITHER the decode NOR the recolor:
 *  1. the color-replacement-free bitmap comes from the shared, per-document,
 *     path-keyed {@link getCachedBitmapByPath} (the raster/metafile cache docx,
 *     pptx and xlsx now share). It content-sniffs the bytes (extension/MIME are
 *     unreliable — sample-10's chart is a standard WMF mislabeled `.emf`),
 *     rasterizing a WMF via the minimal player at a size from `widthPt`/`heightPt`,
 *     returning `null` for a true EMF (or a geometry-less metafile), else
 *     `createImageBitmap`. That `null` is a LEGITIMATE "no drawable output" (not
 *     an error), so we propagate it as `null` and `preloadImages` drops the image
 *     — the existing "missing image" behavior, no crash. A fetch/decode failure
 *     rejects and remains on the renderer/viewer's explicit error path. Every
 *     draw site null-checks the map lookup, matching pptx's
 *     `if (!bitmap) return` and xlsx's "skip if falsy" draw guards.
 *  2. when a clrChange is requested, the make-transparent result is memoized per
 *     (imagePath, colorReplaceFrom) in {@link colorReplacedCacheFor} so the
 *     expensive getImageData/putImageData pass runs once per document.
 *
 * `suppressBoundaryFrame: true` is REQUIRED: docx's former in-tree player ran the
 * window/device-boundary edge suppression unconditionally (to hide sample-10's
 * Fig.1 cosmetic outer frame). Core defaults that heuristic OFF (spec-clean), so
 * docx must opt in here to preserve its current rendering.
 *
 * Exported for unit testing of the lazy-bytes contract.
 */
export async function decodeRaster(
  imagePath: string,
  mimeType: string,
  colorReplaceFrom: string | undefined,
  fetchImage: (path: string, mime: string) => Promise<Blob>,
  widthPt = 0,
  heightPt = 0,
  duotone?: Duotone,
): Promise<ImageBitmap | null> {
  // Base bitmap (no colour replacement): shared, path-keyed, per-document cache.
  const base = await getCachedBitmapByPath(imagePath, mimeType, fetchImage, {
    widthPt,
    heightPt,
    suppressBoundaryFrame: true,
  });
  // A `null` base is a legitimate "no drawable output" (a true EMF or a
  // geometry-less metafile), NOT an error: propagate it so `preloadImages`
  // drops the image and every draw site skips it via its null-check. We return
  // null rather than throw so this expected outcome never travels the exception
  // path. A fetch/decode failure still rejects into the active renderer/viewer
  // error contract; only a stale render suppresses it after the canvas ownership
  // token shows that a newer render has superseded the operation.
  if (!base) return null;
  if (!colorReplaceFrom && !duotone) return base;
  // Second layer: memoize the recolour result per (path, colour, duotone). The
  // recolour reads the SHARED base bitmap and produces a fresh independent raster,
  // so the base is never mutated and stays reusable for other references / draws.
  // clrChange (§20.1.8.11 make-transparent) is applied BEFORE the duotone
  // (§20.1.8.23 luminance ramp): the ramp leaves fully-transparent pixels
  // untouched, so a colour keyed transparent stays transparent under the recolour.
  const cache = colorReplacedCacheFor(fetchImage);
  const key = imageKey(imagePath, colorReplaceFrom, duotone);
  let hit = cache.get(key);
  if (!hit) {
    hit = (async () => {
      let bmp: ImageBitmap = base;
      if (colorReplaceFrom) bmp = await applyColorReplacement(bmp, colorReplaceFrom);
      if (duotone) {
        const { w, h } = imageNaturalSize(bmp);
        if (w > 0 && h > 0) {
          bmp = (await applyDuotone(bmp, duotone, { width: w, height: h })) as ImageBitmap;
        }
      }
      return bmp;
    })();
    // Don't poison the cache if the recolor pass rejects; let the next call retry.
    hit.catch(() => cache.delete(key));
    // A PASS-THROUGH result (duotone-only with a degenerate size or an
    // unavailable pixel pipeline — `applyDuotone` returned the base unchanged)
    // must not be memoized beyond its in-flight window: the resolved value IS
    // the base bitmap, whose lifetime the shared base cache owns (its LRU may
    // evict and GPU-close it later), and a lingering second-layer entry would
    // keep serving the closed bitmap while bypassing the base layer's
    // remove-on-evict → re-decode protection. Same rule as core's
    // getCachedDuotoneBitmapByPath; a fresh recolour raster stays memoized.
    void hit
      .then((bmp) => {
        if (bmp === base) cache.delete(key);
      })
      .catch(() => {});
    cache.set(key, hit);
  }
  return hit;
}

/**
 * Decode every embedded image referenced by the document into a drawable map
 * keyed by `imageKey(imagePath, colorReplaceFrom)`. Bytes are fetched lazily by
 * zip path via `fetchImage`; SVG vector originals decode through the path-keyed
 * `<img>` helper. Returns an empty map when `fetchImage` is absent (no byte
 * source) — draw sites then simply skip.
 *
 * Exported for unit testing of the keying + single-decode-per-key contract.
 */
export async function preloadImages(
  doc: DocxDocumentModel,
  fetchImage: ((path: string, mime: string) => Promise<Blob>) | undefined,
  layoutServices?: LayoutServices,
): Promise<Map<string, DecodedImage>> {
  if (!fetchImage) return new Map();
  const fetch = fetchImage;
  const pairs = collectImagePairs(doc, layoutServices ?? createLayoutServices(doc));
  const entries = await Promise.all(
    pairs.map(async (pair): Promise<[string, DecodedImage] | null> => {
      // Unified svgBlip selection (shared with pptx/xlsx). The decoded image is
      // keyed by the raster `imagePath` regardless of which source we picked, so
      // every draw site finds it via imageKey(imagePath, …) unchanged.
      const dataIsSvg = pair.mimeType === 'image/svg+xml';
      // Shared vector-vs-raster gate (see core preferVectorBlip). `hasCrop` is
      // this format's already-aggregated "any reference to this key is cropped"
      // flag, so it stands in for srcRect presence (`|| null` normalises the
      // undefined case). When true, `blip.svgImagePath` is narrowed to string.
      const blip = { svgImagePath: pair.svgImagePath, srcRect: pair.hasCrop || null };
      // `decodeRaster` may resolve to `null` for a legitimately undrawable
      // metafile (true EMF / geometry-less WMF). That is not an error: omit its
      // map entry. Fetch, SVG+raster fallback, decode, and recolor failures are
      // different outcomes and deliberately reject this preload operation.
      let img: DecodedImage | null;
      if (preferVectorBlip(blip)) {
        // Prefer the vector original (Microsoft `asvg:svgBlip` extension);
        // fall back to the raster on any SVG decode failure. With an
        // `<a:srcRect>` crop (§20.1.8.55) we skip this branch and decode the
        // raster instead, because the crop math (drawImageCropped) needs the
        // bitmap's native pixel grid — an SVG element has none.
        try {
          img = await getCachedSvgImageByPath(blip.svgImagePath, fetch);
        } catch (vectorError) {
          // The raster fallback carries the §20.1.8.23 duotone recolour; an SVG
          // vector original has no readable pixel grid, so it stays un-recoloured.
          const fallback = dataIsSvg
            ? await getCachedSvgImageByPath(pair.imagePath, fetch)
            : await decodeRaster(pair.imagePath, pair.mimeType, pair.colorReplaceFrom, fetch, pair.widthPt, pair.heightPt, pair.duotone);
          // A successful fallback is authoritative. A legitimate null raster,
          // however, cannot erase the vector source's real decode failure: no
          // drawable source remains, and classifying that outcome as merely an
          // unsupported metafile would hide package corruption from onError.
          if (!fallback) throw vectorError;
          img = fallback;
        }
      } else if (dataIsSvg) {
        // svg-only picture (no svgImagePath surfaced — e.g. a non-svgBlip
        // `.svg` part): `createImageBitmap` can't rasterize SVG, so decode
        // through the path-keyed <img>-based SVG path.
        img = await getCachedSvgImageByPath(pair.imagePath, fetch);
      } else {
        img = await decodeRaster(pair.imagePath, pair.mimeType, pair.colorReplaceFrom, fetch, pair.widthPt, pair.heightPt, pair.duotone);
      }
      // Undrawable metafile → explicit unavailable resource at session binding.
      if (!img) return null;
      return [imageKey(pair.imagePath, pair.colorReplaceFrom, pair.duotone), img];
    }),
  );
  return new Map(entries.filter((e): e is [string, DecodedImage] => e !== null));
}

// ===== Main entry =====

/**
 * Per-canvas monotonic render token for the {@link renderDocumentToCanvas}
 * cancellation guard. A WeakMap keyed on the canvas replaces the previous
 * property monkey-patch (`canvas.__docxRenderToken`), so no non-standard field
 * is written onto the caller's canvas and the `as unknown as` cast is gone.
 * WeakMap keys are held weakly, so a discarded canvas is collected normally.
 * (Mirrors the pptx renderSlide guard's renderTokens map.)
 */
const renderTokens = new WeakMap<HTMLCanvasElement | OffscreenCanvas, number>();

/** True when a section flows VERTICALLY (glyphs stack top→bottom, lines advance
 *  across the page). `<w:sectPr><w:textDirection>` uses the TRANSITIONAL
 *  ST_TextDirection enum (ECMA-376 Part 4 §14.11.7; Word writes these, not the
 *  Part 1 §17.18.93 Strict `tb|rl|lr|…` set):
 *    - `tbRl`  (≡ Strict `rl`)  — vertical, lines right→left: standard vertical
 *                                 Japanese; the only value in the samples.
 *    - `tbRlV` (≡ Strict `rlV`) — vertical R→L, non-EA glyphs rotated 90° CW.
 *    - `tbLrV` (≡ Strict `lrV`) — vertical L→R, non-EA glyphs rotated 90° CW.
 *  These three share the +90° page rotation + upright-CJK glyph path (stage-1
 *  approximates the `V` variants' non-EA rotation the same as `tbRl`, which the
 *  glyph path already draws Latin sideways for).
 *
 *    - `btLr`  (≡ Strict `lr`)  — its NOMINAL semantics are bottom-to-top /
 *                                 left-to-right, but Word ground truth (issue #988
 *                                 re-adjudication, correcting the batch-3
 *                                 adjudication ① with raster proof on asymmetric
 *                                 glyphs) shows Word renders a SECTION-level
 *                                 `btLr` as the HORIZONTAL layout rotated +90° CW
 *                                 WHOLESALE: same page frame as `tbRl` (columns
 *                                 right→left, advance top→bottom — neither axis
 *                                 honors the nominal bottom-to-top/left-to-right),
 *                                 but EVERY glyph rides the page rotation — CJK is
 *                                 NOT counter-rotated upright (the dakuten of 「び」
 *                                 lands bottom-right) and vertical punctuation
 *                                 forms are NOT substituted. So `btLr` routes
 *                                 through the same +90° FRAME as `tbRl` while
 *                                 `RenderState.verticalAllRotated` switches the
 *                                 glyph draw to the horizontal branches. (The per-
 *                                 SECTION mixing that fixture also exercises —
 *                                 a `btLr` non-final section beside a horizontal
 *                                 final section — is surfaced per-section: the
 *                                 SectionBreak marker carries `textDirection`,
 *                                 the paginator stamps it in lockstep with the
 *                                 section geometry, and each page rotates by its
 *                                 OWN direction. Issue #1000.)
 *  Two are HORIZONTAL (glyphs upright, lines top→bottom) ⇒ false:
 *    - `lrTb`  (≡ Strict `tb`, the default) — dropped to null by the parser.
 *    - `lrTbV` (≡ Strict `tbV`) — horizontal, EA glyphs rotated 270°; still a
 *                                 horizontal flow, so not this vertical path. */
function isVerticalSection(s: SectionProps): boolean {
  return isVerticalTextDirection(s.textDirection);
}

/** The raw-token predicate behind {@link isVerticalSection}, for section-region
 *  carriers that hold a bare `textDirection` value rather than a
 *  full SectionProps (issue #1000). `null`/`undefined`/unknown ⇒ horizontal. */
function isVerticalTextDirection(td: string | null | undefined): boolean {
  return td === 'tbRl' || td === 'tbRlV' || td === 'tbLrV' || td === 'btLr';
}

/** True for a VERTICAL direction whose glyphs ALL ride the +90° page rotation
 *  (no CJK upright counter-rotation, no vertical punctuation forms, no 縦中横):
 *  section-level `btLr` per the issue #988 re-adjudication (Word GT, raster-
 *  proven — see {@link RenderState.verticalAllRotated}). The tbRl family keeps
 *  the upright-CJK glyph path. Only meaningful when
 *  {@link isVerticalTextDirection} is already true. */
function isAllRotatedVerticalTextDirection(td: string | null | undefined): boolean {
  return td === 'btLr';
}

/** Map a vertical (tbRl) section's PHYSICAL page geometry to the SWAPPED LOGICAL
 *  geometry the horizontal layout engine lays the page out in: logical width =
 *  physical height, and the four margins rotate one quarter-turn so the logical
 *  layout, once the page paint is rotated +90° back into physical space, lands
 *  the margins on the correct physical edges (§17.6.11). With the page transform
 *  `physical = (pageW_phys − logical.y, logical.x)`:
 *    logical.marginLeft  (flow start / column top)  → physical TOP     margin
 *    logical.marginTop   (before the first line)     → physical RIGHT   margin
 *    logical.marginRight (after the last line)        → physical LEFT    margin
 *    logical.marginBottom (flow end / column bottom)  → physical BOTTOM  margin
 *  Non-geometry fields (docGrid, columns, textDirection, …) are preserved so the
 *  logical layout keeps the section's grid pitch, columns and vertical flag. */
function verticalLayoutSection(phys: SectionProps): SectionProps {
  return {
    ...phys,
    ...logicalGeomOf(phys),
  };
}

/** Return a shallow copy of `doc` with its BODY-LEVEL section swapped to the
 *  vertical LOGICAL geometry, so the pagination + layout engine — which reads
 *  `doc.section` — organises the page as a rotated horizontal page. Per-body
 *  SectionBreak `geom`s are NOT swapped here: the paginator resolves each
 *  mid-body section's own logical frame (swapped iff THAT section is vertical)
 *  through `sectionFrameFrom` (issue #1000 per-section text direction), so a
 *  marker's `geom` stays the PHYSICAL geometry the parser emitted. Only invoked
 *  when the body section is vertical; horizontal docs are returned untouched
 *  (referential identity), keeping the horizontal path byte-identical. */
function verticalLayoutDoc(doc: DocxDocumentModel): DocxDocumentModel {
  if (!isVerticalSection(doc.section)) return doc;
  return { ...doc, section: verticalLayoutSection(doc.section) };
}

/** Inverse of {@link verticalLayoutSection}: map a vertical section's SWAPPED
 *  LOGICAL geometry back to its PHYSICAL page geometry. Used to render a vertical
 *  section's header/footer, which — per Word ground truth (issue #988) — stay
 *  HORIZONTAL at the physical top/bottom margins and do NOT rotate with the tbRl
 *  body. Applying `verticalLayoutSection` then this returns the original section
 *  (a quarter-turn each way): physical margin{Top,Right,Bottom,Left} =
 *  logical margin{Left,Top,Right,Bottom}; header/footer distances are preserved. */
function physicalLayoutSection(logical: SectionProps): SectionProps {
  return {
    ...logical,
    ...physicalGeomOf(logical),
  };
}

export async function renderDocumentToCanvas(
  doc: DocxDocumentModel,
  canvas: HTMLCanvasElement | OffscreenCanvas,
  pageIndex: number,
  opts: RenderDocumentOptions = {},
): Promise<void> {
  // Render-pass lease (core acquireBitmapCacheLease): `preloadImages` resolves
  // EVERY image the document references into a non-owning lookup map and the
  // page paint then draws from it synchronously. The shared bitmap cache is
  // LRU-bounded, so a document referencing more images than the cap — or a
  // concurrent render of another page of the same document — would otherwise
  // evict AND GPU-close bitmaps this pass's map still holds before the paint.
  // Under the lease the eviction still removes the cache entry (bounded size;
  // the next pass re-decodes), but the close is deferred until this pass ends,
  // so the paint never draws a closed bitmap.
  const releaseLease = opts.fetchImage ? acquireBitmapCacheLease(opts.fetchImage) : undefined;
  try {
    await renderDocumentToCanvasLeased(doc, canvas, pageIndex, opts);
  } finally {
    releaseLease?.();
  }
}

/** {@link renderDocumentToCanvas}'s body, verbatim; runs under the caller's
 *  render-pass lease. */
async function renderDocumentToCanvasLeased(
  doc: DocxDocumentModel,
  canvas: HTMLCanvasElement | OffscreenCanvas,
  pageIndex: number,
  opts: RenderDocumentOptions = {},
): Promise<void> {
  const layoutServices = opts.layoutServices ?? createLayoutServices(
    doc,
    doc.parseError == null ? {
      measureContext: canvas.getContext('2d') as
        | CanvasRenderingContext2D
        | OffscreenCanvasRenderingContext2D
        | null,
    } : {},
  );
  // Cancellation guard. renderDocumentToCanvas is async (it awaits image decode
  // via preloadImages), so rapid page navigation can start a newer render of the
  // SAME canvas before this one finishes. Both clear the canvas (`canvas.width =
  // …` + a white fillRect) up front and then draw their page AFTER the await —
  // so the clears run first and the draws accumulate, ghosting several pages on
  // top of each other. Stamp a per-canvas token; once a newer render supersedes
  // us, stop at the next await so only the latest render's output survives.
  // (Mirrors the pptx renderSlide guard; the worker path renders each page on a
  // fresh OffscreenCanvas, so the token is a no-op there.)
  const myToken = (renderTokens.get(canvas) ?? 0) + 1;
  renderTokens.set(canvas, myToken);
  const superseded = () => renderTokens.get(canvas) !== myToken;

  const dpr = opts.dpr ?? defaultDpr();
  // getContext before sizing is legal: resizing a canvas after getContext resets
  // its drawing state, and the ctx.scale/fill below run AFTER canvas.width/height.
  // Header/footer measurement and retained body acquisition share this context.
  const ctx = canvas.getContext('2d') as CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
  // ECMA-376 §17.6.20 — for a vertical (tbRl) section the page is laid out in a
  // SWAPPED logical space (logical width = physical page height) and rotated +90°
  // into physical space at paint. `layoutDoc` carries the swapped BODY-level
  // geometry through pagination + per-page section resolution; horizontal docs
  // get `doc` unchanged (referential identity ⇒ byte-identical). Text direction
  // is PER-SECTION (issue #1000), so the `vertical` flag is resolved per PAGE
  // below, after the stamped page frame is merged.
  const resolvedLocalFonts = layoutServices.text.localMetrics;
  const defaultCurrentDateMs = opts.defaultCurrentDateMs ?? Date.now();
  const layoutDoc = verticalLayoutDoc(doc);
  const layoutSettings = resolveDocumentLayoutSettings(layoutDoc);
  const kinsoku = layoutSettings.kinsoku;
  if (!layoutVariantStoreOf(layoutServices)) {
    attachDocumentLayoutVariants({
      model: doc,
      services: layoutServices,
      defaultCurrentDateMs,
      buildLayout: (options) => layoutDocument(doc, layoutServices, options),
    });
  }
  const retainedSelection = selectDocumentLayoutPage(layoutServices, {
    currentDate: opts.currentDate,
    defaultCurrentDateMs,
  }, pageIndex);
  const layoutOptions = retainedSelection.options;
  const retainedBodyLayout = retainedSelection.layout;
  const retainedBodyPage = retainedSelection.page;
  const totalPages = retainedBodyLayout.pages.length;
  // ECMA-376 §17.6.12 — canonical page construction owns the displayed number
  // and format; paint only consumes the selected immutable page metadata.
  const thisPageNumber = {
    displayNumber: retainedBodyPage.pageNumber.displayNumber,
    format: retainedBodyPage.pageNumber.format as NumberFormat,
  };

  // ECMA-376 §17.6.13 / §17.6.11 — page geometry is PER-SECTION. Size THIS page from
  // the section active at its top (resolvePageSection.geom, stamped by the paginator),
  // NOT from the single body-level `doc.section`. `sec` merges the resolved geometry
  // (size + margins + header/footer distances) over the body-level section so the
  // docGrid / columns / sectionStart / even-odd fields keep their body-level values —
  // those already flow per-section through canonical region/docGrid state
  // rails, so only the page-box geometry needs the per-page swap here. For a
  // single-section document `geom` equals `doc.section`, so `sec === doc.section` in
  // value — byte-identical output.
  // `sec` is the LOGICAL section the body/header/footer are laid out in: for a
  // vertical page that is the swapped geometry (the paginator stamps a vertical
  // section's SWAPPED logical geom), for horizontal it equals the physical
  // section. All RenderState geometry below (contentX/W, margins, pageWidth,
  // docGrid) reads `sec`, so the entire layout is expressed in logical
  // coordinates and the canonical region transform maps it to physical space.
  //
  // Issue #1000 — `textDirection` is retained per page in lockstep with page
  // geometry, and `vertical` — which keys the
  // physical canvas box, the +90° body transform, `verticalCJK`, `verticalPhys`
  // and the horizontal header/footer branch below — is derived from the MERGED
  // section, so a vertical non-final section and a horizontal final section
  // each paint in their own orientation.
  const pageTd = retainedBodyPage.section.textDirection;
  const sec: SectionProps = {
    ...layoutDoc.section,
    ...retainedBodyPage.section.geometry,
    textDirection: pageTd,
  };
  const vertical = isVerticalSection(sec);
  const sectionLayout = resolveSectionLayoutContext(layoutSettings, sec);

  // The CANVAS is sized to the PHYSICAL page (visible landscape page for tbRl):
  // physical width = logical height, physical height = logical width. `scale`
  // (px per pt) is isotropic, so the logical layout — whose logical width in px
  // is `sec.pageWidth * scale = physicalHeight * scale = cssHeight` — maps 1:1
  // onto the rotated physical box. For horizontal pages physW/H equal sec's and
  // this is the pre-vertical computation unchanged.
  const physPageWidth = retainedBodyPage.geometry.widthPt;
  const physPageHeight = retainedBodyPage.geometry.heightPt;
  const cssWidth = opts.width ?? physPageWidth * PT_TO_PX;
  const scale = cssWidth / physPageWidth;  // px per pt
  const cssHeight = physPageHeight * scale;

  // Clamp the backing store to browser canvas limits (RB5). A pathological page
  // size (or a large dpr × page size) can exceed the per-axis / total-area cap,
  // at which point the browser silently allocates a smaller-or-empty buffer and
  // the page renders blank. `clampCanvasSize` scales BOTH axes by one factor
  // (≤ 1) so the aspect ratio is kept; we fold that factor into the effective
  // dpr, keep the CSS box at its intended size, and the browser stretches the
  // (slightly lower-res) backing store to fill it — a visible page beats a blank
  // one. `effectiveDpr` is stored on the state so crisp-offset math stays aligned
  // with the real backing-store scale.
  const clamped = clampCanvasSize(cssWidth * dpr, cssHeight * dpr);
  const effectiveDpr = clamped.clamped ? dpr * clamped.scale : dpr;

  canvas.width = clamped.width;
  canvas.height = clamped.height;

  if (isHTMLCanvas(canvas)) {
    canvas.style.width = `${cssWidth}px`;
    canvas.style.height = `${cssHeight}px`;
    if (!canvas.style.display) canvas.style.display = 'block';
  }

  ctx.scale(effectiveDpr, effectiveDpr);
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, cssWidth, cssHeight);

  // RB7 partial degradation: a document whose body part (`word/document.xml`)
  // failed to parse (see the Rust `degraded_document`) carries `parseError` and
  // an empty body. Paint a visible error placeholder page instead of a blank
  // white sheet, and stop. Healthy documents (no parseError) are unaffected. This
  // short-circuits before canonical body paint: a degraded page has no logical
  // flow to project, and the placeholder is laid out in physical space.
  if (doc.parseError != null) {
    await paintRetainedLayoutPage(retainedBodyLayout, 0, canvas, { scale, dpr: effectiveDpr });
    return;
  }

  let images: Map<string, DecodedImage>;
  try {
    images = await preloadImages(doc, opts.fetchImage, layoutServices);
  } catch (error) {
    // A stale render must not report a late decode failure after a newer render
    // has taken ownership of this canvas. The active render still rejects so the
    // viewer routes corruption/fetch failures through its onError contract.
    if (superseded()) return;
    throw error;
  }
  // A newer render of this canvas started while we awaited image decode — stop
  // so we don't paint this (now stale) page over the newer one.
  if (superseded()) return;

  const privateResources = privateResourceLookupOf<CanvasImageSource>(layoutServices);
  const retainedResourceSession = createProductionPaintResourceSession(
    paintResourceRegistryOf(layoutServices),
    (descriptor) => {
      if (descriptor.kind === 'math') {
        return privateResources?.keys.includes(descriptor.resourceKey)
          ? privateResources.resolve(descriptor.resourceKey)
          : unavailablePaintResourceHandle('optional math renderer unavailable');
      }
      if (descriptor.kind === 'image' || descriptor.kind === 'picture-bullet') {
        const image = images.get(imageKey(
          descriptor.partPath,
          descriptor.colorReplaceFrom,
          descriptor.duotone as Duotone | undefined,
        ));
        return image ?? unavailablePaintResourceHandle(
          opts.fetchImage
            ? 'unsupported image format produced no drawable output'
            : 'image byte source unavailable',
        );
      }
      return undefined;
    },
  );
  const retainedResourcePainter = createCanvasPaintResourcePainter(
    retainedResourceSession,
    canonicalCanvasPaintResourceHandlers,
  );

  // ECMA-376 §17.11: map each note id to its 1-based display number so the
  // reference markers (and the in-note footnoteRef placeholder) show the
  // sequential number, not the raw @w:id.
  const footnoteNums = buildNoteNumberMap(
    doc.footnotes,
    noteReferenceIdsInDocumentOrder(doc.body, 'footnote'),
  );
  const endnoteNums = buildNoteNumberMap(
    doc.endnotes,
    noteReferenceIdsInDocumentOrder(doc.body, 'endnote'),
  );
  const noteNumbers = new Map<string, number>();
  for (const [id, n] of footnoteNums) noteNumbers.set(`footnote:${id}`, n);
  for (const [id, n] of endnoteNums) noteNumbers.set(`endnote:${id}`, n);

  // ECMA-376 §17.6.11: the body is inset from each page edge by the margin's MAGNITUDE
  // (a negative margin places the body |margin| inside the edge, overlapping the
  // header/footer — see bodyMarginInsetPt). Identity for the non-negative common case.
  // The header/footer reserves below still use the SIGNED margin (header/footerOverflowPt
  // return 0 for a negative one), so a negative margin reserves nothing yet insets |margin|.
  const bodyTopPt = bodyMarginInsetPt(sec.marginTop);
  const bodyBottomPt = bodyMarginInsetPt(sec.marginBottom);
  const baseState: RenderState = {
    ctx,
    scale,
    // The backing store may have been clamped below `cssSize × dpr`; crisp-offset
    // math must use the SAME effective dpr the ctx was scaled by (see above).
    dpr: effectiveDpr,
    pointToCss: vertical
      ? { a: 0, b: scale, c: -scale, d: 0, e: cssWidth, f: 0 }
      : { a: scale, b: 0, c: 0, d: scale, e: 0, f: 0 },
    contentX: sec.marginLeft * scale,
    contentW: (sec.pageWidth - sec.marginLeft - sec.marginRight) * scale,
    y: bodyTopPt * scale,
    // `pageH` is the LOGICAL page height in px (`sec.pageHeight * scale`). For a
    // horizontal page that equals `cssHeight`; for a vertical page the logical
    // height is the physical WIDTH, so it equals `cssWidth`. Using the logical
    // height keeps the body-flow / footnote / bottom-margin math in the same
    // (logical) coordinate space the page transform maps to physical.
    pageH: sec.pageHeight * scale,
    defaultColor: opts.defaultTextColor ?? '#000000',
    pageIndex,
    totalPages,
    // ECMA-376 §17.6.12 — the current page's displayed number + format (per-section
    // restart / fmt), consumed by a PAGE field in the body, header, or footer.
    displayPageNumber: thisPageNumber.displayNumber,
    pageNumberFormat: thisPageNumber.format,
    images,
    dryRun: false,
    marginLeft: sec.marginLeft,
    marginRight: sec.marginRight,
    // §17.6.11: store the body inset (|margin|), the value the paint pass re-adds as a
    // column's region top (state.marginTop, renderBodyElements); never the raw sign.
    marginTop: bodyTopPt,
    marginBottom: bodyBottomPt,
    pageWidth: sec.pageWidth,
    floats: [],
    floatParaSeq: 0,
    docGrid: toLegacyDocGridContext(sectionLayout),
    layoutSettings,
    sectionLayout,
    storyContext: BODY_STORY_CONTEXT,
    docEastAsian: layoutSettings.documentHasEastAsianText,
    fontFamilyClasses: fontClassesWithPitches(doc.fontFamilyClasses, doc.fontFamilyPitches),
    resolvedLocalFonts,
    layoutServices,
    retainedResourcePainter,
    kinsoku,
    // §17.15.1.25 — automatic tab interval, resolved once and threaded like
    // `kinsoku` so the measure and draw passes agree.
    defaultTabPt: layoutSettings.defaultTabPt,
    characterSpacingControl: layoutSettings.characterSpacingControl,
    useFeLayout: layoutSettings.compat.useFeLayout,
    balanceSingleByteDoubleByteWidth:
      layoutSettings.compat.balanceSingleByteDoubleByteWidth,
    mathDefJc: layoutSettings.mathDefJc,
    onTextRun: opts.onTextRun,
    showTrackChanges: opts.showTrackChanges ?? true,
    // §17.16.4.1 — the instant DATE/TIME fields format against (default real time).
    currentDateMs: layoutOptions.currentDateMs,
    noteNumbers,
    // ECMA-376 §17.6.20 — the frame-level vertical flag. On a tbRl-family page
    // the glyph-draw path counter-rotates upright (CJK) glyphs so they stand up
    // inside the +90°-rotated body region (see the canonical region matrix and
    // `drawVerticalRun`); `verticalAllRotated` below suppresses that for btLr.
    verticalCJK: vertical,
    // §17.6.20 btLr (#988 re-adjudication): every glyph rides the +90° page
    // rotation — the glyph-draw sites take the HORIZONTAL branches instead of
    // the upright-CJK vertical ones (see RenderState.verticalAllRotated).
    verticalAllRotated: vertical && isAllRotatedVerticalTextDirection(pageTd),
    // ECMA-376 §20.4.3.x — physical page geometry for resolving DrawingML anchors
    // against the un-rotated physical page (see `verticalPhys` docs and
    // `resolveAnchorBox`). `sec` here is the LOGICAL (swapped) section, so un-swap
    // it back to physical: physical left/top/right/bottom margin = logical
    // bottom/left/top/right (the inverse of `verticalLayoutSection`). Top/bottom
    // are stored as body insets (`bodyMarginInsetPt`) to match the horizontal
    // path's `marginTop`/`marginBottom` (§17.6.11 text-margin = body edge).
    verticalPhys: vertical
      ? {
          pageWidth: physPageWidth,
          pageHeight: physPageHeight,
          marginLeft: sec.marginBottom,
          marginRight: sec.marginTop,
          marginTop: bodyMarginInsetPt(sec.marginLeft),
          marginBottom: bodyMarginInsetPt(sec.marginRight),
          cssWidthPx: cssWidth,
        }
      : undefined,
  };
  const firstPageOfSection = isFirstSectionOwnedPage(retainedBodyLayout.pages, pageIndex);

  // ECMA-376 §17.6.10 — page borders with zOrder="back" are painted UNDER the body
  // flow (behind intersecting text/objects). Drawn here, before the body.
  const pageBorders = retainedBodyPage.pageBorders;
  if (pageBorders != null) {
    const activePageBorders = pageBorders as PageBorders;
    if (activePageBorders.zOrder === 'back' && pageBorderShownOnPage(activePageBorders, firstPageOfSection)) {
      ctx.save();
      try {
        if (vertical) {
          ctx.translate(cssWidth, 0);
          ctx.rotate(Math.PI / 2);
        }
        drawPageBorders(ctx, activePageBorders, sec, scale);
      } finally {
        ctx.restore();
      }
    }
  }
  ctx.save();
  try {
    const pointScale = effectiveDpr * scale;
    if (typeof ctx.setTransform === 'function') {
      ctx.setTransform(pointScale, 0, 0, pointScale, 0, 0);
    } else {
      ctx.scale(scale, scale);
    }
    paintLayoutPageContent(retainedBodyPage as import('./layout/types.js').LayoutPage, {
      ctx,
      scale,
      dpr: effectiveDpr,
      resources: retainedResourcePainter,
      defaultTextColor: baseState.defaultColor,
      showTrackChanges: baseState.showTrackChanges,
      ...(opts.onTextRun ? { onTextRun: opts.onTextRun } : {}),
    });
  } finally {
    ctx.restore();
  }

  // ECMA-376 §17.6.10 — page borders with zOrder="front" (the default) are painted
  // OVER intersecting text/objects, so draw them LAST (after the whole page flow).
  if (pageBorders != null) {
    const activePageBorders = pageBorders as PageBorders;
    if (activePageBorders.zOrder !== 'back' && pageBorderShownOnPage(activePageBorders, firstPageOfSection)) {
      ctx.save();
      try {
        if (vertical) {
          ctx.translate(cssWidth, 0);
          ctx.rotate(Math.PI / 2);
        }
        drawPageBorders(ctx, activePageBorders, sec, scale);
      } finally {
        ctx.restore();
      }
    }
  }
}

/** Retained default separator leading used by the shared note story layout. */
const FOOTNOTE_SEPARATOR_GAP_PT = 6;

/** ECMA-376 §17.10.1 — an empty header/footer set (no default/first/even). Used
 *  when constructing a measure-only document service. */
const EMPTY_HEADERS_FOOTERS: HeadersFooters = { default: null, first: null, even: null };

/** Build default sequential numbering from first-reference order. Unreferenced
 *  note-part entries do not consume a displayed number (§17.18.22/.34). */
function buildNoteNumberMap(
  notes: DocNote[] | undefined,
  referenceIds: readonly string[],
): Map<string, number> {
  const m = new Map<string, number>();
  if (!notes) return m;
  const available = new Set(notes.map((note) => note.id));
  referenceIds.forEach((id) => {
    if (available.has(id) && !m.has(id)) m.set(id, m.size + 1);
  });
  return m;
}

/** Index footnotes by id for content lookup. */
function indexNotes(notes: DocNote[] | undefined): Map<string, DocNote> {
  const m = new Map<string, DocNote>();
  if (!notes) return m;
  for (const n of notes) m.set(n.id, n);
  return m;
}

// Preserve the historical renderer re-export without retaining the removed raw
// page-element carrier.
export type { ColumnGeom } from './types';

// `computeColumns` is the shared pure implementation imported from
// layout-context.ts and re-exported above for existing renderer callers.
/** The document's default body font size in pt, used to size line-number glyphs so
 *  they share the body baseline grid. Resolved from the first body paragraph's
 *  `defaultFontSize` (which the parser folds from docDefaults + the style chain),
 *  falling back to 10pt (the ECMA-376 docDefaults sz absent value). */
function docDefaultFontSizePt(doc: DocxDocumentModel): number {
  for (const el of doc.body) {
    if (el.type === 'paragraph') {
      const p = el as unknown as DocParagraph;
      if (typeof p.defaultFontSize === 'number') return p.defaultFontSize;
      for (const run of p.runs) {
        if (run.type === 'text') return (run as unknown as DocxTextRun).fontSize;
      }
    }
  }
  return 10;
}

/** ECMA-376 §17.6.10 `@w:display` (§17.18.62) — page borders are evaluated only
 * on pages owned by the section, including its first such physical page. */
function pageBorderShownOnPage(pb: PageBorders, firstSectionPage: boolean): boolean {
  switch (pb.display) {
    case 'firstPage':
      return firstSectionPage;
    case 'notFirstPage':
      return !firstSectionPage;
    default: // "allPages" and any unknown value
      return true;
  }
}

/** ECMA-376 §17.6.10 — draw a section's page borders as a rectangle inset from the
 *  page edge (`offsetFrom="page"`) or the text margin (`offsetFrom="text"`, the
 *  default). Each edge's `space` (pt) is the inset from the reference; `sz`→width
 *  and `val`→style reuse the shared border-line drawing (single/double/dashed/…).
 *  Art borders (§17.18.2 decorative-image styles) are unsupported — such a `val`
 *  yields no drawable dash/line and is skipped. */
function drawPageBorders(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  pb: PageBorders,
  sec: SectionProps,
  scale: number,
): void {
  // Reference edges (pt): the page box, or the text-margin box.
  const fromText = pb.offsetFrom === 'text';
  const refLeftPt = fromText ? sec.marginLeft : 0;
  const refRightPt = fromText ? sec.pageWidth - sec.marginRight : sec.pageWidth;
  const refTopPt = fromText ? bodyMarginInsetPt(sec.marginTop) : 0;
  const refBottomPt = fromText ? sec.pageHeight - bodyMarginInsetPt(sec.marginBottom) : sec.pageHeight;

  // Each edge is inset from its reference by that edge's `space` (pt), TOWARD the
  // page interior: the top border moves DOWN, bottom UP, left RIGHT, right LEFT.
  const asSpec = (e: PageBorderEdge): BorderSpec => ({ width: e.width, color: e.color ?? null, style: e.style });
  const topY = (refTopPt + (pb.top?.space ?? 0)) * scale;
  const bottomY = (refBottomPt - (pb.bottom?.space ?? 0)) * scale;
  const leftX = (refLeftPt + (pb.left?.space ?? 0)) * scale;
  const rightX = (refRightPt - (pb.right?.space ?? 0)) * scale;

  // The four sides span between the two perpendicular inset lines so corners meet.
  if (pb.top) drawBorderLine(ctx, leftX, topY, rightX, topY, asSpec(pb.top), scale, 1);
  if (pb.bottom) drawBorderLine(ctx, leftX, bottomY, rightX, bottomY, asSpec(pb.bottom), scale, 1);
  if (pb.left) drawBorderLine(ctx, leftX, topY, leftX, bottomY, asSpec(pb.left), scale, 1);
  if (pb.right) drawBorderLine(ctx, rightX, topY, rightX, bottomY, asSpec(pb.right), scale, 1);
}

export function layoutDocument(
  doc: DocxDocumentModel,
  services: LayoutServices = createLayoutServices(doc),
  options: LayoutOptions = normalizeLayoutOptions(undefined, Date.now()),
): RetainedDocumentLayout {
  const normalizedDoc = normalizeInternalDocumentModel(doc).document;
  return paginateBody(createBodyLayoutInput(normalizedDoc), services, options);
}

function buildMeasureState(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  section: SectionProps,
  fontFamilyClasses: Record<string, string> = {},
  layoutSettings: DocumentLayoutSettings,
  resolvedLocalFonts: Readonly<Record<string, ResolvedLocalFontMetric>> = {},
  layoutServices?: LayoutServices,
  layoutOptions?: LayoutOptions,
): RenderState {
  const sectionLayout = resolveSectionLayoutContext(layoutSettings, section);
  // Story acquisition may omit services. Acquisition itself no longer
  // has an optional text authority: construct the same A2 service at this stable
  // state boundary so every downstream buildSegments/layoutLines call records
  // authoritative grapheme geometry. Production document entry points already
  // pass their document-scoped service and therefore keep their exact font
  // inventory/resource session.
  const effectiveLayoutServices = layoutServices ?? createLayoutServices({
    section,
    body: [],
    headers: EMPTY_HEADERS_FOOTERS,
    footers: EMPTY_HEADERS_FOOTERS,
    fontFamilyClasses,
  }, {
    measureContext: ctx,
    localMetrics: resolvedLocalFonts,
  });
  return {
    ctx,
    scale: 1,
    dpr: 1,
    // Mirror the PAINT pass seed (renderDocumentToCanvas: `contentX =
    // sec.marginLeft × scale`; scale is 1 here). contentX/contentW carry the
    // current text column, and §20.4.3.4 `relativeFrom="column"` anchors
    // resolve against them (xContainer). Seeding 0 made the MEASURE pass place
    // body-level column anchors a full marginLeft LEFT of where the paint pass
    // draws them, so floats entered/left the wrap band only during pagination
    // and paragraphs split differently from the painted layout (PR #844 review
    // F1; pinned by paginate-column-anchor.test.ts).
    contentX: section.marginLeft,
    contentW: section.pageWidth - section.marginLeft - section.marginRight,
    y: 0,
    pageH: section.pageHeight,
    defaultColor: '#000000',
    pageIndex: 0,
    totalPages: fieldAcquisitionContextOf(effectiveLayoutServices).totalPages,
    images: new Map(),
    dryRun: true,
    marginLeft: section.marginLeft,
    marginRight: section.marginRight,
    // §17.6.11: the measure state's marginTop is the BODY-LEVEL body inset (|margin|).
    // Canonical per-section regions no longer read this field directly: the split
    // functions derive the region top from the threaded `tagSectionGeom` closure
    // (`bodyMarginInsetPt(tagSectionGeom().marginTop)`), matching pushTagged's
    // `bodyTopPt()` per-section convention. This body-level value is only the
    // single-section-equivalent fallback (identical when there is one section) and
    // still seeds contentW/pageH below. Never the raw sign. Identity for non-negative.
    marginTop: bodyMarginInsetPt(section.marginTop),
    marginBottom: bodyMarginInsetPt(section.marginBottom),
    pageWidth: section.pageWidth,
    floats: [],
    floatParaSeq: 0,
    docGrid: toLegacyDocGridContext(sectionLayout),
    layoutSettings,
    sectionLayout,
    storyContext: BODY_STORY_CONTEXT,
    docEastAsian: layoutSettings.documentHasEastAsianText,
    fontFamilyClasses,
    resolvedLocalFonts,
    layoutServices: effectiveLayoutServices,
    retainedTableAcquisition: {
      layoutServices: (state) => state.layoutServices,
      tableFormat: tableFormatInput,
      resolveColumns: resolveColumnWidths,
      createCellState: (state, contentWidthPt, cell) => ({
        ...withTableCellStory(state),
        contentX: 0,
        contentW: contentWidthPt,
        y: 0,
        containerShading: cell.background ?? state.containerShading,
        floats: [],
        floatParaSeq: 0,
        pageAnchorPrescanned: new Set<DocParagraph>(),
      }),
      acquireParagraph: (
        cellState,
        paragraph,
        paragraphWidthPt,
        paragraphPath,
        flowDomainId,
        paragraphBorderEdges,
        inheritedAuthority,
        sourceRef,
      ) => {
        const source = sourceRef ?? {
          story: 'body' as const,
          storyInstance: 'body',
          path: [...paragraphPath],
        };
        const publicRuns = paragraph.runs.filter((run, runIndex) =>
          publicAnchorBridge(run, source, runIndex) !== null);
        if (publicRuns.length > 0) {
          // Hand-built compatibility runs have no parser anchor/host acquisition
          // facts, so their paragraph-top projection stays outside the parser
          // fixed point until the public bridge is removed.
          registerAnchorFloats(
            { ...paragraph, runs: publicRuns },
            cellState,
            cellState.y,
          );
        }
        const context = resolveStateParagraphLayoutContext(cellState, paragraph);
        const layout = acquireRegisteredParagraph(
          cellState,
          paragraphAcquisitionInput(paragraph, source),
          {
            id: `${source.story}:${source.storyInstance}:${source.path.join('.')}`,
            source,
            flowDomainId,
            ordinaryFlow: true,
            context,
            placement: {
              startYPt: cellState.y,
              paragraphXPt: 0,
              availableWidthPt: paragraphWidthPt,
              maximumYPt: cellState.pageH,
              suppressSpaceBefore: true,
            },
            measurer: {
              context: cellState.ctx,
              fontFamilyClasses: cellState.fontFamilyClasses,
            },
            environment: paragraphMeasurementEnvironment(cellState),
            exclusions: paragraphWrapExclusions(cellState.floats, flowDomainId),
            anchorCollisions: paragraphAnchorCollisions(cellState.floats),
            anchorCellBounds: {
              xPt: 0,
              yPt: 0,
              widthPt: paragraphWidthPt,
              heightPt: cellState.pageH / cellState.scale,
            },
            containerShading: cellState.containerShading,
            ...(paragraphBorderEdges ? { paragraphBorderEdges } : {}),
            trailingExtentPt: Math.max(
              context.spaceAfterPt,
              paragraphBorderEdges?.bottom === 'none'
                ? 0
                : bottomBorderExtentPt(paragraph.borders),
            ),
            continuesFromPrevious: false,
            anchorFrames: paragraphAnchorReferenceFrames(cellState),
            acquireCompleteStory: cellState.acquireCompleteTextBoxStory,
          },
          inheritedAuthority,
        ).layout;
        if (paragraph.spaceBefore === 0) return layout;
        return Object.freeze({
          ...layout,
          flowBounds: Object.freeze({
            ...layout.flowBounds,
            heightPt: layout.flowBounds.heightPt + paragraph.spaceBefore,
          }),
          advancePt: layout.advancePt + paragraph.spaceBefore,
          spacing: Object.freeze({
            ...layout.spacing,
            beforePt: paragraph.spaceBefore,
          }),
        });
      },
      registerFloatingTable: (state, request) => {
        const scale = state.scale;
        const usesTextX = !request.positioning.horzSpecified
          || (request.positioning.horzAnchor !== 'page'
            && request.positioning.horzAnchor !== 'margin');
        const usesTextY = request.positioning.vertAnchor !== 'page'
          && request.positioning.vertAnchor !== 'margin';
        // Page/margin coordinates are not final until the containing table is
        // paginated. Registering them in this cell-local acquisition state would
        // reserve a different rectangle from the later page-local paint box.
        if (!usesTextX || !usesTextY) return null;
        const pageHeightPt = state.pageH / scale;
        const textFrame = {
          xPt: state.contentX / scale,
          yPt: state.y / scale,
          widthPt: state.contentW / scale,
          heightPt: request.child.advancePt,
        };
        const box = resolveFloatingTableBoxPt(
          request.positioning,
          {
            page: {
              xPt: 0,
              yPt: 0,
              widthPt: state.pageWidth,
              heightPt: pageHeightPt,
            },
            margin: {
              xPt: state.marginLeft,
              yPt: state.marginTop,
              widthPt: Math.max(0, state.pageWidth - state.marginLeft - state.marginRight),
              heightPt: Math.max(0, pageHeightPt - state.marginTop - state.marginBottom),
            },
            text: textFrame,
          },
          request.child.columnWidthsPt.reduce((sum, width) => sum + width, 0),
          request.child.advancePt,
        );
        const registered = pushFloatRect(state, {
          x: box.x * scale,
          y: box.y * scale,
          w: box.w * scale,
          h: box.h * scale,
          dl: request.positioning.leftFromTextPt * scale,
          dr: request.positioning.rightFromTextPt * scale,
          dt: request.positioning.topFromTextPt * scale,
          db: request.positioning.bottomFromTextPt * scale,
          kind: 'table',
          mode: 'square',
          side: 'bothSides',
          imageKey: '',
          drawn: true,
          paraId: state.floatParaSeq++,
          avoidOverlap: true,
          allowOverlap: request.overlap !== 'never',
        });
        return Object.freeze({
          xPt: registered.imageX / scale - textFrame.xPt,
          yPt: registered.imageY / scale - textFrame.yPt,
        });
      },
      advanceState: (state, advancePt) => {
        state.y += advancePt;
      },
    },
    retainedTablesBySourceIndex: new Map<number, RetainedTableRecord>(),
    currentDateMs: layoutOptions?.currentDateMs,
    kinsoku: layoutSettings.kinsoku,
    defaultTabPt: layoutSettings.defaultTabPt,
    characterSpacingControl: layoutSettings.characterSpacingControl,
    useFeLayout: layoutSettings.compat.useFeLayout,
    balanceSingleByteDoubleByteWidth:
      layoutSettings.compat.balanceSingleByteDoubleByteWidth,
    showTrackChanges: false,
    // ECMA-376 §17.6.20 + §20.4.3.x (issue #988 ②, Codex review F1): for a
    // vertical (tbRl) section — `section` is the SWAPPED logical geometry — the
    // measure pass must resolve DrawingML anchors against the same PHYSICAL
    // page the paint pass uses (`resolveAnchorBox`/`resolveShapeBox` key their
    // physical branch on `verticalPhys`), otherwise a wrapped shape's exclusion
    // band is reserved at the raw logical rectangle during pagination while the
    // paint wraps around the physical projection — diverging page assignment.
    // Mirrors the paint-state seed (renderDocumentToCanvas), un-swapping via
    // physicalLayoutSection; `cssWidthPx` at the paginator's scale 1 is the
    // physical page width in pt. `verticalCJK` stays UNSET: the measure pass
    // keeps its horizontal glyph metrics (only anchor geometry re-frames).
    // Seeded from the section this measure state is BUILT from (the body-level
    // body-level one); a direction-mixed document then re-seeds it per
    // section via its retained acquisition location (issue #1000), so a
    // mid-body section's anchors resolve against ITS OWN physical frame.
    get verticalCJK() {
      return isVerticalTextDirection(this.sectionLayout.textDirection);
    },
    get verticalAllRotated() {
      return isVerticalTextDirection(this.sectionLayout.textDirection)
        && isAllRotatedVerticalTextDirection(this.sectionLayout.textDirection);
    },
    verticalPhys: isVerticalSection(section)
      ? (() => {
          const phys = physicalLayoutSection(section);
          return {
            pageWidth: phys.pageWidth,
            pageHeight: phys.pageHeight,
            marginLeft: phys.marginLeft,
            marginRight: phys.marginRight,
            marginTop: bodyMarginInsetPt(phys.marginTop),
            marginBottom: bodyMarginInsetPt(phys.marginBottom),
            cssWidthPx: phys.pageWidth,
          };
        })()
      : undefined,
  };
}

function createConcreteBodyLayoutKernel(
  doc: DocxDocumentModel,
  measureContext: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D | null,
  resolvedLocalFonts: Readonly<Record<string, ResolvedLocalFontMetric>>,
): BodyLayoutKernel {
  const ordinaryAcquisitionInputForAdjacentGroup = (
    group: ReturnType<typeof combineAdjacentTableLayoutInputs>,
  ): TableLayoutInput => {
    const noEdges = Object.freeze({
      top: null, right: null, bottom: null, left: null, insideH: null, insideV: null,
    });
    return Object.freeze({
      kind: 'table',
      id: group.id,
      source: group.source,
      flowDomainId: group.flowDomainId,
      ordinaryFlow: true,
      alignment: group.alignment,
      indentPt: group.indentPt,
      bidiVisual: group.bidiVisual,
      columnWidthsPt: group.columnWidthsPt,
      columnWidthKeys: group.columnWidthKeys,
      borders: noEdges,
      rows: Object.freeze(group.rows.map((row) => Object.freeze({
        ...row,
        // §17.4.37 gives every authored member table ownership of its own outer
        // border layer; the union grid carries that folded layer per source row.
        exceptionBorders: row.sourceTableEdges,
      }))),
    });
  };
  const sourceElement = (
    source: SourceRef,
  ): Extract<BodyElement, { type: 'paragraph' | 'table' }> => {
    if (source.story !== 'body' || source.storyInstance !== 'body' || source.path.length !== 1) {
      throw new Error('Body acquisition requires a top-level body source');
    }
    const element = doc.body[source.path[0]!];
    if (!element || (element.type !== 'paragraph' && element.type !== 'table')) {
      throw new Error(`Body source does not identify a flow block: ${source.path.join('.')}`);
    }
    return element;
  };
  const nestedSourceElement = (source: SourceRef): BodyElement | CellElement => {
    if (source.story !== 'body' || source.storyInstance !== 'body'
      || source.path.length === 0 || (source.path.length - 1) % 3 !== 0) {
      throw new Error('Nested body acquisition requires a canonical source path');
    }
    let element: BodyElement | CellElement | undefined = doc.body[source.path[0]!];
    for (let offset = 1; offset < source.path.length; offset += 3) {
      if (!element || element.type !== 'table') {
        throw new Error(`Nested body source leaves table ownership: ${source.path.join('.')}`);
      }
      element = element.rows[source.path[offset]!]
        ?.cells[source.path[offset + 1]!]
        ?.content[source.path[offset + 2]!];
    }
    if (!element || (element.type !== 'paragraph' && element.type !== 'table')) {
      throw new Error(`Nested body source does not identify a flow block: ${source.path.join('.')}`);
    }
    return element;
  };
  /** Body acquisition stays at the kernel adapter because it resolves legacy
   * renderer state into retained paragraph inputs; layout-owned projections
   * below it receive only immutable structural values. */
  const acquireBodyParagraphAtLocation = (
    state: RenderState,
    paragraph: DocParagraph,
    source: SourceRef,
    location: BodyAcquisitionLocation,
    availableInlineExtentPt: number,
    suppressSpaceBefore: boolean,
    continuation: BodyParagraphAcquisitionInput['continuation'] = Object.freeze({
      boundary: null,
    }),
    retainedAnchorCollisions?: readonly DrawingMLCollisionEntryPt[],
  ) => {
    const edges = bodyParagraphBorderEdgesFor(paragraph) ?? {
      top: 'top' as const,
      bottom: 'bottom' as const,
    };
    const context = resolveBodyParagraphLayoutContext(state, paragraph);
    return acquireParagraphResult(
      paragraphAcquisitionInput(paragraph, source),
      {
        id: `${source.story}:${source.storyInstance}:${source.path.join('.')}`,
        source,
        flowDomainId: location.flowDomainId,
        ordinaryFlow: true,
        context,
        placement: {
          startYPt: state.y,
          paragraphXPt: location.availableBounds.xPt,
          availableWidthPt: availableInlineExtentPt,
          maximumYPt: state.pageH,
          suppressSpaceBefore,
        },
        measurer: { context: state.ctx, fontFamilyClasses: state.fontFamilyClasses },
        environment: paragraphMeasurementEnvironment(state),
        exclusions: paragraphWrapExclusions(state.floats, location.flowDomainId),
        anchorCollisions: retainedAnchorCollisions
          ?? paragraphAnchorCollisions(state.floats),
        containerShading: state.containerShading,
        paragraphBorderEdges: edges,
        trailingExtentPt: Math.max(
          context.spaceAfterPt,
          edges.bottom === 'none' ? 0 : bottomBorderExtentPt(paragraph.borders),
        ),
        continuesFromPrevious: continuation.boundary !== null,
        ...(continuation.sourceRangeStart === undefined ? {} : {
          sourceRangeStart: continuation.sourceRangeStart,
        }),
        anchorFrames: paragraphAnchorReferenceFrames(state),
        acquireCompleteStory: state.acquireCompleteTextBoxStory,
      },
      continuation.boundary === null ? undefined : {
        boundary: continuation.boundary,
        ...(continuation.uniformRubyAdvancePt === undefined ? {} : {
          uniformRubyAdvancePt: continuation.uniformRubyAdvancePt,
        }),
      },
    );
  };
  return Object.freeze({
    openBodyLayoutSession(
      input: import('./layout/body-layout-kernel.js').BodyLayoutSessionInput,
      services: LayoutServices,
      options: LayoutOptions,
    ) {
      if (!measureContext) throw new Error('Body layout acquisition requires a measurement context');
      const physicalSection: SectionProps = {
        ...doc.section,
        ...input.section.geometry,
        textDirection: input.section.textDirection,
        vAlign: input.section.verticalAlignment,
      };
      const section = isVerticalTextDirection(physicalSection.textDirection)
        ? verticalLayoutSection(physicalSection)
        : physicalSection;
      const state = buildMeasureState(
        measureContext,
        section,
        fontClassesWithPitches(doc.fontFamilyClasses, doc.fontFamilyPitches),
        resolveDocumentLayoutSettings(doc),
        resolvedLocalFonts,
        services,
        options,
      );
      const footnotesById = indexNotes(doc.footnotes ?? []);
      state.noteNumbers = new Map([
        ...[...buildNoteNumberMap(
          doc.footnotes,
          noteReferenceIdsInDocumentOrder(doc.body, 'footnote'),
        )].map(
          ([id, number]) => [`footnote:${id}`, number] as const,
        ),
        ...[...buildNoteNumberMap(
          doc.endnotes,
          noteReferenceIdsInDocumentOrder(doc.body, 'endnote'),
        )].map(
          ([id, number]) => [`endnote:${id}`, number] as const,
        ),
      ]);
      const stories = new Map<string, HeaderFooter>();
      createBodySectionIndex(bodySectionIndexInput(doc)).occurrences.forEach((occurrence) => {
        const prefix = occurrence.markerBodyIndex === null
          ? null
          : `section:${occurrence.markerBodyIndex}`;
        for (const kind of ['default', 'first', 'even'] as const) {
          const storyInstance = prefix === null ? kind : `${prefix}:${kind}`;
          const header = occurrence.headers[kind];
          const footer = occurrence.footers[kind];
          if (header) stories.set(`header:${storyInstance}`, header);
          if (footer) stories.set(`footer:${storyInstance}`, footer);
        }
      });
      let location = input.initialLocation;
      const pageRegistryFlowDomainId = (pageIndex: number) => `body:page:${pageIndex}:registry`;
      let floatRegistry: FloatRegistrySnapshotPt = Object.freeze({
        coordinateSpace: 'logical-page-points' as const,
        flowDomainId: pageRegistryFlowDomainId(location.pageIndex),
        entries: Object.freeze([]) as readonly FloatRegistryEntryPt[],
        nextParagraphId: 0,
      });
      let drawingCollisionRegistry: DrawingMLCollisionRegistrySnapshotPt =
        createDrawingMLCollisionRegistry(
          pageRegistryFlowDomainId(location.pageIndex),
          'logical-page-points',
        );
      const applyLocationTo = (target: RenderState, next: BodyAcquisitionLocation) => {
        const geometry = next.section.geometry;
        target.sectionLayout = next.section as SectionLayoutContext;
        target.docGrid = toLegacyDocGridContext(next.section as SectionLayoutContext);
        target.pageIndex = next.pageIndex;
        const page = fieldAcquisitionContextOf(services).resolveDestinationPage?.(next.pageIndex);
        target.displayPageNumber = page?.displayPageNumber ?? next.pageIndex + 1;
        target.pageNumberFormat = page?.pageNumberFormat ?? target.pageNumberFormat;
        target.pageWidth = geometry.pageWidth;
        target.pageH = geometry.pageHeight;
        target.marginLeft = geometry.marginLeft;
        target.marginRight = geometry.marginRight;
        target.marginTop = bodyMarginInsetPt(geometry.marginTop);
        target.marginBottom = bodyMarginInsetPt(geometry.marginBottom);
        target.contentX = next.availableBounds.xPt;
        target.contentW = next.availableBounds.widthPt;
        target.y = next.cursorPt.yPt;
      };
      const applyLocation = (next: BodyAcquisitionLocation) => {
        location = next;
        applyLocationTo(state, next);
      };
      applyLocation(location);
      const publicParagraphFloatAcquisition = (
        paragraph: DocParagraph,
        source: SourceRef,
        candidate: RenderState,
        onlyOccurrenceIds?: ReadonlySet<string>,
        paragraphId = floatRegistry.nextParagraphId,
      ): readonly FloatRegistryEntryPt[] => {
        const committedOccurrenceIds = new Set(floatRegistry.entries.map((entry) => entry.occurrenceId));
        const publicRuns = paragraph.runs.flatMap((run, runIndex) => {
          if (run.type !== 'shape' && run.type !== 'image' && run.type !== 'chart') return [];
          const bridge = publicAnchorBridge(run, source, runIndex);
          if (!bridge
            || (onlyOccurrenceIds && !onlyOccurrenceIds.has(bridge.occurrenceId))
            || committedOccurrenceIds.has(bridge.occurrenceId)
            || (bridge.pageOwned && candidate.pageAnchorPrescanned?.has(paragraph))) return [];
          return [{
            run,
            occurrenceId: bridge.occurrenceId,
          }];
        });
        if (publicRuns.length === 0) return Object.freeze([]);
        const baseFloatCount = candidate.floats.length;
        registerAnchorFloats(
          { ...paragraph, runs: publicRuns.map(({ run }) => run) },
          candidate,
          candidate.y,
        );
        const registered = candidate.floats.slice(baseFloatCount);
        if (registered.length !== publicRuns.length) {
          throw new Error('Public paragraph anchor acquisition did not retain every wrap float');
        }
        const scale = candidate.scale;
        return Object.freeze(registered.map((float, index): FloatRegistryEntryPt => {
          const occurrenceId = publicRuns[index]!.occurrenceId;
          return Object.freeze({
            kind: 'shape',
            occurrenceId,
            exclusionId: occurrenceId,
            paragraphId,
            bounds: Object.freeze({
              xPt: float.imageX / scale,
              yPt: float.imageY / scale,
              widthPt: float.imageW / scale,
              heightPt: float.imageH / scale,
            }),
            exclusionBounds: Object.freeze({
              xPt: float.xLeft / scale,
              yPt: float.yTop / scale,
              widthPt: (float.xRight - float.xLeft) / scale,
              heightPt: (float.yBottom - float.yTop) / scale,
            }),
            wrap: publicRuns[index]!.run.wrapMode as NonNullable<FloatRegistryEntryPt['wrap']>,
            wrapSide: float.side,
            wrapDistances: Object.freeze({
              topPt: float.distTop / scale,
              rightPt: float.distRight / scale,
              bottomPt: float.distBottom / scale,
              leftPt: float.distLeft / scale,
            }),
            ...(float.wrapPolygon ? { wrapPolygon: Object.freeze([...float.wrapPolygon]) } : {}),
          });
        }));
      };
      const retainedParagraphFloatEntries = (
        layout: ParagraphLayout,
      ): readonly FloatRegistryEntryPt[] => {
        const hostFrames = new Map((layout.anchorFrames ?? []).flatMap((frame) => {
          if (frame.status !== 'resolved') return [];
          const isHostAxis = (axis: typeof frame.axes.horizontal) => axis.status === 'resolved'
            && (axis.referenceFrame === 'paragraph'
              || axis.referenceFrame === 'line'
              || axis.referenceFrame === 'character');
          return isHostAxis(frame.axes.horizontal) || isHostAxis(frame.axes.vertical)
            ? [[frame.occurrenceId, frame] as const]
            : [];
        }));
        if (hostFrames.size === 0) return Object.freeze([]);
        const exclusions = new Map(layout.exclusions.flatMap((exclusion) =>
          exclusion.anchorOccurrenceId
            ? [[exclusion.anchorOccurrenceId, exclusion] as const]
            : []));
        return Object.freeze((layout.anchorCollisions ?? []).flatMap(
          (collision): FloatRegistryEntryPt[] => {
          const frame = hostFrames.get(collision.occurrenceId);
          if (!frame) return [];
          if (frame.geometry.wrap.kind === 'none') return [];
          const exclusion = exclusions.get(collision.occurrenceId);
          if (!exclusion) {
            throw new Error(`Wrapped anchor omitted exclusion geometry: ${collision.occurrenceId}`);
          }
          return [Object.freeze({
            kind: 'shape' as const,
            occurrenceId: collision.occurrenceId,
            exclusionId: collision.occurrenceId,
            paragraphId: floatRegistry.nextParagraphId,
            bounds: collision.bounds,
            exclusionBounds: exclusion.bounds,
            horizontalOwnership: collision.horizontalOwnership,
            verticalOwnership: collision.verticalOwnership,
            wrap: frame.geometry.wrap.kind,
            wrapSide: frame.geometry.wrap.side,
            wrapDistances: frame.geometry.wrap.distances,
            ...(frame.geometry.wrap.polygon
              ? { wrapPolygon: frame.geometry.wrap.polygon.points }
              : {}),
          })];
        }));
      };
      const reacquireTableBlock = (
        request: PageDependentTableBlockRequest,
      ): ParagraphLayout | TableLayout => {
        if (request.acquired.kind !== 'paragraph') return request.acquired;
        const source = nestedSourceElement(request.acquired.source);
        if (source.type !== 'paragraph') {
          throw new Error('Table paragraph re-acquisition source kind mismatch');
        }
        const scale = state.scale;
        const candidate: RenderState = {
          ...withTableCellStory(state),
          contentX: 0,
          contentW: request.acquired.flowBounds.widthPt * scale,
          y: request.acquired.flowBounds.yPt,
          floats: (request.floatingTableExclusions ?? []).map((bounds, index): FloatRect => ({
            kind: 'table', mode: 'square',
            imageKey: `${TRANSIENT_TABLE_FINAL_FRAME_EXCLUSION_PREFIX}${index}`,
            imageX: bounds.xPt * scale, imageY: bounds.yPt * scale,
            imageW: bounds.widthPt * scale, imageH: bounds.heightPt * scale,
            xLeft: bounds.xPt * scale,
            xRight: (bounds.xPt + bounds.widthPt) * scale,
            yTop: bounds.yPt * scale,
            yBottom: (bounds.yPt + bounds.heightPt) * scale,
            side: 'bothSides', distLeft: 0, distRight: 0, distTop: 0, distBottom: 0,
            paraId: index, drawn: true,
          })),
          floatParaSeq: request.floatingTableExclusions?.length ?? 0,
          pageAnchorPrescanned: new Set<DocParagraph>(),
        };
        const inheritedAuthority =
          inheritedParagraphAuthorityForReacquisition(request.acquired);
        const tableAcquisition = state.retainedTableAcquisition;
        if (!tableAcquisition) {
          throw new Error('Table paragraph re-acquisition requires retained dependencies');
        }
        return tableAcquisition.acquireParagraph(
          candidate,
          source,
          request.acquired.flowBounds.widthPt,
          request.acquired.source.path,
          request.acquired.flowDomainId,
          undefined,
          inheritedAuthority,
        );
      };
      const endnotesById = indexNotes(doc.endnotes ?? []);
      const storyLayoutCache = new Map<string, StoryLayout>();
      const storyRoot = (source: SourceRef): readonly BodyElement[] => {
        if (source.path.length !== 0) {
          throw new Error('Story acquisition requires a story-root source');
        }
        if (source.story === 'header' || source.story === 'footer') {
          const story = stories.get(`${source.story}:${source.storyInstance}`);
          if (!story) throw new Error(`Unknown ${source.story} story source`);
          return story.body;
        }
        if (source.story === 'footnote' || source.story === 'endnote') {
          const note = (source.story === 'footnote' ? footnotesById : endnotesById)
            .get(source.storyInstance);
          if (!note) throw new Error(`Unknown ${source.story} story source`);
          return note.content;
        }
        throw new Error(`Unsupported shared story source: ${source.story}`);
      };
      const storyElement = (
        root: readonly (BodyElement | CompleteTextBoxBlockInput)[],
        source: SourceRef,
      ): Extract<BodyElement, { type: 'paragraph' | 'table' }> => {
        if (source.path.length === 0 || (source.path.length - 1) % 3 !== 0) {
          throw new Error('Story block acquisition requires a canonical source path');
        }
        type TraversableStoryElement = Readonly<{
          type: string;
          rows?: readonly Readonly<{
            cells: readonly Readonly<{
              content: readonly TraversableStoryElement[];
            }>[];
          }>[];
        }>;
        let element = root[source.path[0]!] as TraversableStoryElement | undefined;
        for (let offset = 1; offset < source.path.length; offset += 3) {
          if (!element || element.type !== 'table') {
            throw new Error(`Story source leaves table ownership: ${source.path.join('.')}`);
          }
          element = element.rows?.[source.path[offset]!]
            ?.cells[source.path[offset + 1]!]
            ?.content[source.path[offset + 2]!];
        }
        if (!element || (element.type !== 'paragraph' && element.type !== 'table')) {
          throw new Error(`Story source does not identify a flow block: ${source.path.join('.')}`);
        }
        return element as unknown as Extract<BodyElement, { type: 'paragraph' | 'table' }>;
      };
      const acquireStoryLayout = (
        request: import('./layout/body-layout-kernel.js').StoryLayoutAcquisitionInput,
        explicitRoot?: readonly CompleteTextBoxBlockInput[],
      ): StoryLayout => {
        const cacheKey = JSON.stringify({
          source: request.source,
          pageIndex: request.pageIndex,
          section: request.section,
          container: request.container,
        });
        const cached = storyLayoutCache.get(cacheKey);
        if (cached) return cached;
        const root = explicitRoot ?? storyRoot(request.source);
        const noteReferenceNumber = request.source.story === 'footnote'
          || request.source.story === 'endnote'
          ? state.noteNumbers?.get(
              `${request.source.story}:${request.source.storyInstance}`,
            )
          : undefined;
        const fieldContext = fieldAcquisitionContextOf(services);
        const pageFieldContext = fieldContext.resolveDestinationPage?.(request.pageIndex);
        const storyVertical = isVerticalTextDirection(request.section.textDirection);
        const candidate: RenderState = {
          ...state,
          sectionLayout: request.section as SectionLayoutContext,
          docGrid: toLegacyDocGridContext(request.section as SectionLayoutContext),
          pageIndex: request.pageIndex,
          totalPages: fieldContext.totalPages,
          displayPageNumber: pageFieldContext?.displayPageNumber ?? request.pageIndex + 1,
          pageNumberFormat: pageFieldContext?.pageNumberFormat ?? state.pageNumberFormat,
          pageWidth: request.section.geometry.pageWidth,
          pageH: request.container.capacity === 'unbounded'
            ? Number.MAX_SAFE_INTEGER
            : request.section.geometry.pageHeight,
          marginLeft: request.section.geometry.marginLeft,
          marginRight: request.section.geometry.marginRight,
          marginTop: bodyMarginInsetPt(request.section.geometry.marginTop),
          marginBottom: bodyMarginInsetPt(request.section.geometry.marginBottom),
          contentX: request.container.bounds.xPt,
          contentW: request.container.bounds.widthPt,
          y: request.container.bounds.yPt,
          floats: [],
          floatParaSeq: 0,
          retainedTablesBySourceIndex: new Map(),
          pageAnchorPrescanned: new Set<DocParagraph>(),
          noteReferenceNumber,
          verticalCJK: storyVertical,
          verticalAllRotated: storyVertical
            && isAllRotatedVerticalTextDirection(request.section.textDirection),
          ...(storyVertical ? {} : { verticalPhys: undefined }),
          storyContext: {
            story: request.source.story,
            containers: [],
            lineNumberingEligible: false,
          },
        };
        preRegisterPageFloats(root as readonly BodyElement[], 0, candidate);
        const storyServices = createLayoutServicesRuntimeView(services);
        candidate.layoutServices = storyServices;
        const blockInputs: StoryBlockInput[] = root.flatMap((element, index): StoryBlockInput[] => {
          const source: SourceRef = {
            story: request.source.story,
            storyInstance: request.source.storyInstance,
            path: [index],
          };
          if (element.type === 'unsupportedTextBoxBlock') {
            return [{
              type: 'unsupportedTextBoxBlock',
              qName: element.qName,
              sourcePath: element.sourcePath,
            }];
          }
          if (element.type === 'paragraph') return [{ kind: 'paragraph', source }];
          if (element.type !== 'table') {
            throw new Error(`Unsupported ${request.source.story} story block: ${element.type}`);
          }
          const dependencies = candidate.retainedTableAcquisition;
          if (!dependencies) throw new Error('Story table acquisition requires retained dependencies');
          const table = element as unknown as DocTable;
          const columns = resolveColumnWidths(
            table,
            request.container.bounds.widthPt,
            candidate,
          );
          return [acquireRetainedTable(
            table,
            columns,
            request.container.bounds.widthPt,
            candidate,
            source,
            dependencies,
          ).input];
        });
        let previousParagraph: DocParagraph | null = null;
        const algorithms: BlockLayoutAlgorithms = {
          layoutParagraph(block, placement) {
            const paragraph = storyElement(root, block.source);
            if (paragraph.type !== 'paragraph') throw new Error('Story paragraph source kind mismatch');
            const sourceIndex = block.source.path[0]!;
            const previous = sourceIndex > 0 && root[sourceIndex - 1]?.type === 'paragraph'
              ? root[sourceIndex - 1] as DocParagraph
              : null;
            const next = root[sourceIndex + 1]?.type === 'paragraph'
              ? root[sourceIndex + 1] as DocParagraph
              : null;
            const previousAfterPt = previousParagraph?.spaceAfter ?? 0;
            const spacing = paragraphGapAdjustment(
              previousParagraph,
              paragraph,
              previousAfterPt,
              paragraph.spaceBefore,
            );
            const startYPt = Math.max(
              placement.container.bounds.yPt,
              placement.cursor.yPt - spacing.overlap,
            );
            candidate.y = startYPt;
            candidate.contentX = placement.container.bounds.xPt;
            candidate.contentW = placement.container.bounds.widthPt;
            const publicRuns = paragraph.runs.filter((run, runIndex) =>
              publicAnchorBridge(run, block.source, runIndex) !== null);
            if (publicRuns.length > 0) {
              registerAnchorFloats(
                { ...paragraph, runs: publicRuns },
                candidate,
                candidate.y,
              );
            }
            const context = resolveStateParagraphLayoutContext(candidate, paragraph);
            const borderEdges = resolveParagraphBorderEdges(previous, paragraph, next);
            const result = acquireRegisteredParagraph(
              candidate,
              paragraphAcquisitionInput(paragraph, block.source),
              {
                id: `${block.source.story}:${block.source.storyInstance}:${block.source.path.join('.')}`,
                source: block.source,
                flowDomainId: placement.container.id,
                ordinaryFlow: true,
                context,
                placement: {
                  startYPt,
                  paragraphXPt: placement.container.bounds.xPt,
                  availableWidthPt: placement.container.bounds.widthPt,
                  maximumYPt: placement.availableBounds.yPt + placement.availableBounds.heightPt,
                  suppressSpaceBefore: spacing.suppressBefore,
                },
                measurer: {
                  context: candidate.ctx,
                  fontFamilyClasses: candidate.fontFamilyClasses,
                },
                environment: paragraphMeasurementEnvironment(candidate),
                exclusions: paragraphWrapExclusions(candidate.floats, placement.container.id),
                anchorCollisions: paragraphAnchorCollisions(candidate.floats),
                containerShading: candidate.containerShading,
                paragraphBorderEdges: borderEdges,
                trailingExtentPt: Math.max(
                  context.spaceAfterPt,
                  borderEdges.bottom === 'none' ? 0 : bottomBorderExtentPt(paragraph.borders),
                ),
                continuesFromPrevious: false,
                anchorFrames: paragraphAnchorReferenceFrames(candidate),
                acquireCompleteStory: candidate.acquireCompleteTextBoxStory,
              },
            );
            previousParagraph = paragraph;
            const nextCursor = {
              xPt: placement.cursor.xPt,
              yPt: startYPt + result.layout.advancePt,
            };
            candidate.y = nextCursor.yPt;
            return { layout: result.layout, nextCursor };
          },
          layoutTable(block, placement) {
            previousParagraph = null;
            const normalizedInput: TableLayoutInput = {
              ...block,
              flowDomainId: placement.container.id,
            };
            const result = layoutRetainedTableInput(normalizedInput, placement, storyServices);
            candidate.y = result.nextCursor.yPt;
            return result;
          },
        };
        attachStoryBlockLayoutAlgorithms(storyServices, algorithms);
        const acquired = layoutSharedStory({
          source: request.source,
          container: request.container,
          blocks: Object.freeze(blockInputs),
        }, storyServices);
        const retained = Object.freeze({
          ...acquired,
          blocks: Object.freeze(acquired.blocks.map((block, index) => {
            if (block.kind !== 'paragraph' && block.kind !== 'table') {
              throw new Error(`Shared story emitted unsupported node: ${block.kind}`);
            }
            return projectBodyOccurrence(block, {
              occurrenceId: `${request.container.id}:block:${index}`,
              destination: {
                coordinateSpace: 'logical-page-points',
                flowDomainId: request.container.id,
                translation: { xPt: 0, yPt: 0 },
              },
            });
          })),
        });
        storyLayoutCache.set(cacheKey, retained);
        return retained;
      };
      const acquireCompleteTextBoxStory: CompleteTextBoxStoryAcquirer = (request) =>
        acquireStoryLayout({
          source: request.source,
          pageIndex: state.pageIndex,
          section: state.sectionLayout,
          container: request.container,
        }, request.blocks);
      state.acquireCompleteTextBoxStory = acquireCompleteTextBoxStory;
      const session: BodyLayoutSession = {
        hasPaginationFields: paginatedFlowHasPaginationDependentFields(
          doc.body,
          doc.footnotes ?? [],
          [
            ...[...stories.values()].map((story) => story.body),
            ...(doc.endnotes ?? []).map((note) => note.content),
          ],
        ),
        measureParagraph(request: BodyParagraphAcquisitionInput) {
          applyLocation(request.location);
          const paragraph = sourceElement(request.input.source);
          if (paragraph.type !== 'paragraph') throw new Error('Paragraph source kind mismatch');
          if (paragraph.framePr) {
            if (request.continuation.boundary !== null) {
              throw new Error('Body frame acquisition cannot continue across flow regions');
            }
            let acquiredGroup: ReturnType<typeof acquireRetainedFrameGroup> | undefined;
            const frameGroup = bodyFrameGroupFor(paragraph);
            const box = resolveFrameBox(
              paragraph,
              state,
              frameAnchorLineHeightPx(doc.body, paragraph, state),
              (acquired) => { acquiredGroup = acquired; },
            );
            if (!acquiredGroup) throw new Error('Body frame acquisition omitted its retained group');
            const member = acquiredGroup.members.find((candidate) => candidate.paragraph === paragraph);
            if (!member) throw new Error('Body frame acquisition omitted its retained member');
            if (!frameGroup) throw new Error('Body frame acquisition omitted its adjacency group');
            const absoluteVertical = paragraph.framePr.vAnchor === 'page'
              || paragraph.framePr.vAnchor === 'margin';
            const frameOccurrenceId = box.exclusionId
              ?? `frame:${request.input.source.path.join(':')}`;
            const frameEntry: FloatRegistryEntryPt = Object.freeze({
              kind: 'frame',
              occurrenceId: frameOccurrenceId,
              exclusionId: frameOccurrenceId,
              paragraphId: floatRegistry.nextParagraphId,
              bounds: Object.freeze({
                xPt: box.x / state.scale,
                yPt: box.y / state.scale,
                widthPt: box.w / state.scale,
                heightPt: box.h / state.scale,
              }),
              exclusionBounds: Object.freeze({
                xPt: box.exLeft / state.scale,
                yPt: box.exTop / state.scale,
                widthPt: (box.exRight - box.exLeft) / state.scale,
                heightPt: (box.exBottom - box.exTop) / state.scale,
              }),
            });
            return Object.freeze({
              layout: member.fragment,
              blockExtentPt: 0,
              lineEndBoundaries: Object.freeze([]),
              placement: Object.freeze({
                coordinateSpace: 'logical-body' as const,
                xPt: member.fragment.flowBounds.xPt,
                yPt: member.fragment.flowBounds.yPt,
                sectionFlowOwnership: absoluteVertical ? 'page' as const : 'host-flow' as const,
              }),
              ...(paragraph === frameGroup.owner ? {
                // §17.3.1.11 makes identical adjacent framePr paragraphs one frame,
                // so page admission belongs to the owner before any member is painted.
                retainedFootnoteReferenceIds: Object.freeze([...new Set(
                  acquiredGroup.members.flatMap((candidate) =>
                    footnoteIdsInRetainedSlice(candidate.fragment)),
                )]),
              } : {}),
              ...(!absoluteVertical ? {
                relocationBlockExtentPt: Math.max(
                  0,
                  (box.y + box.h) / state.scale - request.location.cursorPt.yPt,
                ),
              } : {}),
              ...(box.registerExclusion === false ? {} : {
                flowRegistryDelta: Object.freeze({
                  floats: floatingTableRegistryDelta(
                    floatRegistry,
                    Object.freeze([frameEntry]),
                    floatRegistry.nextParagraphId + 1,
                  ),
                }),
              }),
            });
          }
          const candidate: RenderState = {
            ...state,
            floats: [...state.floats],
            pageAnchorPrescanned: new Set(state.pageAnchorPrescanned),
          };
          applyLocationTo(candidate, request.location);
          const publicFloats = request.continuation.boundary === null
            ? publicParagraphFloatAcquisition(paragraph, request.input.source, candidate)
            : Object.freeze([]);
          const acquired = acquireBodyParagraphAtLocation(
            candidate,
            paragraph,
            request.input.source,
            request.location,
            request.availableInlineExtentPt,
            request.suppressSpaceBefore,
            request.continuation,
            drawingCollisionRegistry.entries,
          );
          const { measured, layout } = acquired;
          const allBoundaries = measured.lines.map((line) => {
            const boundary = line.layout.consumedEnd;
            if (!boundary) throw new Error('Measured line omitted its source boundary');
            return boundary;
          });
          const retainedFloats = retainedParagraphFloatEntries(layout);
          const floatEntries = Object.freeze([...publicFloats, ...retainedFloats]);
          // This accepted-collision path is intentionally parser-owned. Hand-built
          // public-model anchors still use the compatibility float bridge (and a
          // public wrapNone run therefore has no collision entry) until the Series
          // B/C bridge removal migrates those runs to the retained OOXML contract.
          const collisionEntries = ownedParagraphAnchorCollisions(layout);
          return Object.freeze({
            layout,
            blockExtentPt: layout.advancePt,
            lineEndBoundaries: Object.freeze(allBoundaries),
            ...(measured.markOnly
              ? { markBelowBaselinePt: measured.lastLineBelowBaselinePt }
              : {}),
            ...(measured.uniformRubyAdvancePt == null
              ? {}
              : { uniformRubyAdvancePt: measured.uniformRubyAdvancePt }),
            ...(floatEntries.length === 0 && collisionEntries.length === 0
              ? {}
              : {
                  flowRegistryDelta: Object.freeze({
                    ...(floatEntries.length === 0 ? {} : {
                      floats: floatingTableRegistryDelta(
                        floatRegistry,
                        floatEntries,
                        floatRegistry.nextParagraphId + floatEntries.length,
                      ),
                    }),
                    ...(collisionEntries.length === 0 ? {} : {
                      drawingCollisions: drawingMLCollisionRegistryDelta(
                        drawingCollisionRegistry,
                        collisionEntries,
                      ),
                    }),
                  }),
                }),
          });
        },
        measureTable(request: BodyTableAcquisitionInput) {
          applyLocation(request.location);
          if (request.input.kind === 'adjacent-table-group') {
            if (request.cursor && request.cursor.kind !== 'adjacent-table-group') {
              throw new Error('Adjacent table group acquisition received an ordinary table cursor');
            }
            const records = request.input.tables.map((tableInput) => {
              const table = sourceElement(tableInput.source);
              if (table.type !== 'table') throw new Error('Table source kind mismatch');
              const sourceIndex = tableInput.source.path[0]!;
              computeTablePtLayout(state, table, request.availableInlineExtentPt, sourceIndex);
              return retainedTableRecord(state, sourceIndex).acquisition;
            });
            const combinedInput = ordinaryAcquisitionInputForAdjacentGroup(
              combineAdjacentTableLayoutInputs(
                request.input.logicalSequenceId,
                records.map((record) => record.input),
              ),
            );
            const placement = {
              container: {
                id: request.location.flowDomainId,
                kind: 'body' as const,
                bounds: {
                  xPt: 0, yPt: 0,
                  widthPt: request.availableInlineExtentPt,
                  heightPt: request.freshPageBlockExtentPt,
                },
              },
              cursor: { xPt: 0, yPt: 0 },
              availableBounds: {
                xPt: 0, yPt: 0,
                widthPt: request.availableInlineExtentPt,
                heightPt: request.freshPageBlockExtentPt,
              },
            };
            const combinedLayout = layoutRetainedTableInput(
              combinedInput,
              placement,
              services,
            ).layout;
            const nestedById: Record<string, RetainedTableAcquisition> = {};
            records.forEach((record) => Object.entries(record.nestedById).forEach(([id, nested]) => {
              if (nestedById[id] && nestedById[id] !== nested) {
                throw new Error(`Adjacent table group has duplicate nested table id: ${id}`);
              }
              nestedById[id] = nested;
            }));
            const combined: RetainedTableAcquisition = Object.freeze({
              input: combinedInput,
              layout: combinedLayout,
              nestedById: Object.freeze(nestedById),
              floatingTables: Object.freeze(records.flatMap((record) => record.floatingTables)),
            });
            const groupCursor: import('./layout/body-layout-kernel.js').AdjacentTableGroupCursor = request.cursor?.cursor ?? Object.freeze({
              tableIndex: 0,
              sourceRowIndex: 0,
            });
            const rowsBefore = request.input.tables
              .slice(0, groupCursor.tableIndex)
              .reduce((sum, tableInput) => sum + (tableInput.rowCount ?? 0), 0);
            const globalRowIndex = rowsBefore + groupCursor.sourceRowIndex;
            const cursor = groupCursor.tableCursor ?? Object.freeze({
              ...startTableFragmentCursor(),
              rowIndex: globalRowIndex,
            });
            if (cursor.rowIndex !== globalRowIndex) {
              throw new Error('Adjacent-table group and table-fragment cursors disagree');
            }
            const result = takeTableFragment(combined, cursor, {
              availableHeightPt: request.availableBlockExtentPt,
              freshPageHeightPt: request.freshPageBlockExtentPt,
              placement,
              services,
              compatibility: 'word',
              page: {
                physicalPageIndex: request.location.pageIndex,
                displayPageNumber: request.location.pageIndex + 1,
                occurrenceId: `${combinedInput.id}:body:${request.location.pageIndex}`,
              },
            });
            if (!result.fragment || result.requiresFreshPage) {
              return Object.freeze({
                layout: combined.layout,
                blockExtentPt: 0,
                nextCursor: Object.freeze({
                  kind: 'adjacent-table-group' as const,
                  cursor: groupCursor,
                }),
                requiresFreshFlowRegion: true,
              });
            }
            const nextGroupCursor = result.nextCursor
              ? (() => {
                  let tableIndex = 0;
                  let firstRow = 0;
                  while (tableIndex < request.input.tables.length) {
                    const rowCount = request.input.tables[tableIndex]!.rowCount ?? 0;
                    if (result.nextCursor!.rowIndex < firstRow + rowCount) break;
                    firstRow += rowCount;
                    tableIndex += 1;
                  }
                  if (tableIndex >= request.input.tables.length) return null;
                  return Object.freeze({
                    tableIndex,
                    sourceRowIndex: result.nextCursor!.rowIndex - firstRow,
                    tableCursor: result.nextCursor!,
                  });
                })()
              : null;
            return Object.freeze({
              layout: result.fragment,
              blockExtentPt: result.fragment.advancePt,
              nextCursor: nextGroupCursor
                ? Object.freeze({ kind: 'adjacent-table-group' as const, cursor: nextGroupCursor })
                : null,
              ...(result.floatingTableRegistryDelta
                ? {
                    flowRegistryDelta: Object.freeze({
                      floats: result.floatingTableRegistryDelta,
                    }),
                  }
                : {}),
            });
          }
          const table = sourceElement(request.input.source);
          if (table.type !== 'table') throw new Error('Table source kind mismatch');
          const sourceIndex = request.input.source.path[0]!;
          computeTablePtLayout(state, table, request.availableInlineExtentPt, sourceIndex);
          const retained = retainedTableRecord(state, sourceIndex).acquisition;
          if (request.cursor && request.cursor.kind !== 'table') {
            throw new Error('Ordinary table acquisition received an adjacent-group cursor');
          }
          const cursor = request.cursor?.cursor ?? startTableFragmentCursor();
          const pageHeightPt = state.pageH / state.scale;
          const authoredPositioning = tableFormatInput(table).positioning;
          if (authoredPositioning) {
            const positioning = request.cursor?.kind === 'table'
              && request.cursor.floatingContinuationFrame === 'fresh-text'
              ? Object.freeze({ ...authoredPositioning, vertAnchor: 'text', yPt: 0, yAlign: undefined })
              : authoredPositioning;
            const tableWidthPt = retained.layout.columnWidthsPt.reduce((sum, width) => sum + width, 0);
            const frames = Object.freeze({
              page: Object.freeze({ xPt: 0, yPt: 0, widthPt: state.pageWidth, heightPt: pageHeightPt }),
              margin: Object.freeze({
                xPt: state.marginLeft, yPt: state.marginTop,
                widthPt: Math.max(0, state.pageWidth - state.marginLeft - state.marginRight),
                heightPt: Math.max(0, pageHeightPt - state.marginTop - state.marginBottom),
              }),
              text: Object.freeze({
                xPt: request.location.cursorPt.xPt,
                yPt: request.location.cursorPt.yPt,
                widthPt: request.availableInlineExtentPt,
                heightPt: retained.layout.advancePt,
              }),
            });
            const raw = resolveFloatingTableBoxPt(
              positioning,
              frames,
              tableWidthPt,
              retained.layout.advancePt,
            );
            const pageAnchoredCollision = request.cursor?.kind !== 'table'
              && (positioning.vertAnchor === 'page' || positioning.vertAnchor === 'margin')
              && floatRegistry.entries.some((entry) => entry.kind === 'table'
                && raw.x + raw.w - entry.exclusionBounds.xPt > FLOAT_OVERLAP_EPS
                && entry.exclusionBounds.xPt + entry.exclusionBounds.widthPt - raw.x
                  > FLOAT_OVERLAP_EPS
                && raw.y + raw.h - entry.exclusionBounds.yPt > FLOAT_OVERLAP_EPS
                && entry.exclusionBounds.yPt + entry.exclusionBounds.heightPt - raw.y
                  > FLOAT_OVERLAP_EPS);
            if (pageAnchoredCollision) {
              // Word defers an absolute page/margin table whose authored band is
              // already owned by a table on this page; the fresh page preserves
              // the authored anchor instead of converting it to a text continuation.
              return Object.freeze({
                layout: retained.layout,
                blockExtentPt: 0,
                nextCursor: Object.freeze({
                  kind: 'table' as const,
                  cursor,
                  floatingContinuationFrame: 'authored' as const,
                }),
                requiresFreshFlowRegion: true,
              });
            }
            const absoluteAnchorMustSplit = (positioning.vertAnchor === 'page'
              || positioning.vertAnchor === 'margin')
              && retained.layout.advancePt > request.freshPageBlockExtentPt;
            const admissionBlockEndPt = absoluteAnchorMustSplit
              ? request.location.availableBounds.yPt
                + request.location.availableBounds.heightPt
              : positioning.vertAnchor === 'page'
                ? frames.page.yPt + frames.page.heightPt
                : positioning.vertAnchor === 'margin'
                  ? frames.margin.yPt + frames.margin.heightPt
                  : request.location.availableBounds.yPt
                    + request.location.availableBounds.heightPt;
            const freshAdmissionHeightPt = absoluteAnchorMustSplit
              ? request.freshPageBlockExtentPt
              : positioning.vertAnchor === 'page'
              ? frames.page.heightPt
              : positioning.vertAnchor === 'margin'
                ? frames.margin.heightPt
                : request.freshPageBlockExtentPt;
            let parentFrame = Object.freeze({ xPt: raw.x, yPt: raw.y });
            let result: ReturnType<typeof takeTableFragment>;
            let resolved: ReturnType<typeof resolveFloatingTablePlacementInTransaction>;
            let nestedEntries: readonly FloatRegistryEntryPt[];
            const visitedTransactions = new Set<string>();
            while (true) {
              const availableHeightPt = Math.max(0, admissionBlockEndPt - parentFrame.yPt);
              result = takeTableFragment(retained, cursor, {
                availableHeightPt,
                freshPageHeightPt: freshAdmissionHeightPt,
                placement: {
                  container: {
                    id: `${request.location.flowDomainId}:floating-table`, kind: 'body',
                    bounds: {
                      xPt: 0, yPt: 0,
                      widthPt: request.availableInlineExtentPt, heightPt: availableHeightPt,
                    },
                  },
                  cursor: { xPt: 0, yPt: 0 },
                  availableBounds: {
                    xPt: 0, yPt: 0,
                    widthPt: request.availableInlineExtentPt, heightPt: availableHeightPt,
                  },
                },
                services,
                compatibility: 'word',
                oversizedRowPolicy: 'atomic',
                page: {
                  physicalPageIndex: request.location.pageIndex,
                  displayPageNumber: state.displayPageNumber ?? request.location.pageIndex + 1,
                  occurrenceId: `${retained.input.id}:fitting-outer:${request.location.pageIndex}:${cursor.rowIndex}:${cursor.rowFragmentIndex}`,
                },
                floatingTableFrames: {
                  page: frames.page,
                  margin: frames.margin,
                  column: frames.text,
                },
                floatingTableRegistry: floatRegistry,
                finalPlacementTranslationPt: parentFrame,
                reacquirePageDependentBlock: reacquireTableBlock,
              });
              if (!result.fragment || result.requiresFreshPage) {
                return Object.freeze({
                  layout: retained.layout,
                  blockExtentPt: 0,
                  nextCursor: Object.freeze({
                    kind: 'table' as const,
                    cursor,
                    floatingContinuationFrame: 'fresh-text' as const,
                  }),
                  requiresFreshFlowRegion: true,
                });
              }
              const sourcePlacement: FloatingTablePlacementLayout = Object.freeze({
                kind: 'floating-table-placement',
                occurrenceId: `${retained.input.id}:root:${request.location.pageIndex}:${cursor.rowIndex}:${cursor.rowFragmentIndex}`,
                ownership: 'source',
                physicalPageIndex: request.location.pageIndex,
                displayPageNumber: state.displayPageNumber ?? request.location.pageIndex + 1,
                hostCellId: request.location.flowDomainId,
                sourceBlockIndex: request.input.source.path[0]!,
                anchorBlockIndex: request.input.source.path[0]!,
                tableId: result.fragment.id,
                overlap: table.overlap === 'never' ? 'never' : 'overlap',
                positioning,
                anchorBounds: frames.text,
                child: result.fragment,
              });
              nestedEntries = result.floatingTableRegistryDelta?.entries ?? [];
              const nestedNextParagraphId = result.floatingTableRegistryDelta?.nextParagraphId
                ?? floatRegistry.nextParagraphId;
              resolved = resolveFloatingTablePlacementInTransaction(
                sourcePlacement,
                frames,
                beginFloatingTablePlacementTransaction(
                  floatRegistry.entries,
                  nestedNextParagraphId,
                  floatRegistry.coordinateSpace,
                  floatRegistry.flowDomainId,
                ),
              );
              if (resolved.placement.xPt === parentFrame.xPt
                && resolved.placement.yPt === parentFrame.yPt) break;
              const fingerprint = JSON.stringify({
                parentFrame,
                fragment: result.fragment,
                nestedEntries,
                resolvedBounds: resolved.placement.bounds,
              });
              if (visitedTransactions.has(fingerprint)) {
                throw new Error('Floating table parent/child transaction did not converge');
              }
              visitedTransactions.add(fingerprint);
              parentFrame = Object.freeze({
                xPt: resolved.placement.xPt,
                yPt: resolved.placement.yPt,
              });
            }
            const isFloatingContinuation = request.cursor?.kind === 'table'
              && request.cursor.floatingContinuationFrame !== undefined;
            const admittedBlockEndPt = request.location.availableBounds.yPt
              + request.location.availableBounds.heightPt;
            const hostFlowPlacements = [
              ...result.fragment.resolvedFloatingTables ?? [],
              resolved.placement,
            ].filter((placement) => placement.source.positioning.vertAnchor === 'text');
            if (!isFloatingContinuation && hostFlowPlacements.some((placement) => (
              placement.exclusionBounds.yPt + placement.exclusionBounds.heightPt
                > admittedBlockEndPt
            ))) {
              return Object.freeze({
                layout: result.fragment,
                blockExtentPt: 0,
                nextCursor: Object.freeze({
                  kind: 'table' as const,
                  cursor,
                  floatingContinuationFrame: 'fresh-text' as const,
                }),
                requiresFreshFlowRegion: true,
              });
            }
            return Object.freeze({
              layout: result.fragment,
              blockExtentPt: 0,
              nextCursor: result.nextCursor
                ? Object.freeze({
                    kind: 'table' as const,
                    cursor: result.nextCursor,
                    floatingContinuationFrame: 'fresh-text' as const,
                  })
                : null,
              flowRegistryDelta: Object.freeze({
                floats: floatingTableRegistryDelta(
                  floatRegistry,
                  Object.freeze([...nestedEntries, ...resolved.transaction.delta]),
                  resolved.transaction.nextParagraphId,
                ),
              }),
              placement: Object.freeze({
                coordinateSpace: 'logical-body' as const,
                xPt: resolved.placement.xPt,
                yPt: resolved.placement.yPt,
                sectionFlowOwnership: positioning.vertAnchor === 'page'
                  || positioning.vertAnchor === 'margin'
                  ? 'page' as const
                  : 'host-flow' as const,
              }),
            });
          }
          if (state.verticalPhys && !effectiveTablePositioning(table)) {
            if (request.cursor) {
              throw new Error('An upright physical table must remain atomic');
            }
            const physical = state.verticalPhys;
            const tableWidthPt = retained.layout.columnWidthsPt.reduce((sum, width) => sum + width, 0);
            if (tableWidthPt > request.availableBlockExtentPt
              && request.availableBlockExtentPt < request.freshPageBlockExtentPt) {
              return Object.freeze({
                layout: retained.layout,
                blockExtentPt: 0,
                nextCursor: Object.freeze({ kind: 'table' as const, cursor }),
                requiresFreshFlowRegion: true,
              });
            }
            const physicalLeftPt = physical.cssWidthPx / state.scale
              - request.location.cursorPt.yPt - tableWidthPt;
            const physicalTopPt = request.location.cursorPt.xPt;
            const physicalBandHeightPt = Math.max(
              retained.layout.advancePt,
              physical.pageHeight - physical.marginTop - physical.marginBottom,
            );
            const flowDomainId = `upright-physical-page:${request.location.pageIndex}`;
            const upright = takeTableFragment(retained, startTableFragmentCursor(), {
              availableHeightPt: physicalBandHeightPt,
              freshPageHeightPt: physicalBandHeightPt,
              placement: {
                container: {
                  id: flowDomainId, kind: 'body',
                  bounds: { xPt: 0, yPt: 0, widthPt: tableWidthPt, heightPt: physicalBandHeightPt },
                },
                cursor: { xPt: 0, yPt: 0 },
                availableBounds: {
                  xPt: 0, yPt: 0, widthPt: tableWidthPt, heightPt: physicalBandHeightPt,
                },
              },
              services,
              compatibility: 'word',
              oversizedRowPolicy: 'atomic',
              page: {
                physicalPageIndex: request.location.pageIndex,
                displayPageNumber: state.displayPageNumber ?? request.location.pageIndex + 1,
                occurrenceId: `${retained.input.id}:upright-page:${request.location.pageIndex}`,
              },
              floatingTableFrames: {
                page: { xPt: 0, yPt: 0, widthPt: physical.pageWidth, heightPt: physical.pageHeight },
                margin: {
                  xPt: physical.marginLeft, yPt: physical.marginTop,
                  widthPt: Math.max(0, physical.pageWidth - physical.marginLeft - physical.marginRight),
                  heightPt: Math.max(0, physical.pageHeight - physical.marginTop - physical.marginBottom),
                },
                column: {
                  xPt: physical.marginLeft, yPt: physical.marginTop,
                  widthPt: Math.max(0, physical.pageWidth - physical.marginLeft - physical.marginRight),
                  heightPt: Math.max(0, physical.pageHeight - physical.marginTop - physical.marginBottom),
                },
              },
              floatingTableRegistry: Object.freeze({
                coordinateSpace: 'upright-physical-page-points' as const,
                flowDomainId,
                entries: Object.freeze([]),
                nextParagraphId: 0,
              }),
              finalPlacementTranslationPt: { xPt: physicalLeftPt, yPt: physicalTopPt },
              reacquirePageDependentBlock: reacquireTableBlock,
            });
            if (!upright.fragment || upright.nextCursor || upright.requiresFreshPage) {
              throw new Error('Upright table final-frame layout must remain atomic');
            }
            return Object.freeze({
              layout: upright.fragment,
              blockExtentPt: tableWidthPt,
              nextCursor: null,
              placement: Object.freeze({
                coordinateSpace: 'upright-physical' as const,
                xPt: physicalLeftPt + upright.fragment.flowBounds.xPt,
                yPt: physicalTopPt + upright.fragment.flowBounds.yPt,
                sectionFlowOwnership: 'host-flow' as const,
              }),
            });
          }
          const result = takeTableFragment(retained, cursor, {
            availableHeightPt: request.availableBlockExtentPt,
            freshPageHeightPt: request.freshPageBlockExtentPt,
            placement: {
              container: {
                id: request.location.flowDomainId,
                kind: 'body',
                bounds: {
                  xPt: 0, yPt: 0,
                  widthPt: request.availableInlineExtentPt,
                  heightPt: request.availableBlockExtentPt,
                },
              },
              cursor: { xPt: 0, yPt: 0 },
              availableBounds: {
                xPt: 0, yPt: 0,
                widthPt: request.availableInlineExtentPt,
                heightPt: request.availableBlockExtentPt,
              },
            },
            services,
            compatibility: 'word',
            page: {
              physicalPageIndex: request.location.pageIndex,
              displayPageNumber: request.location.pageIndex + 1,
              occurrenceId: `${retained.input.id}:body:${request.location.pageIndex}`,
            },
            floatingTableFrames: {
              page: { xPt: 0, yPt: 0, widthPt: state.pageWidth, heightPt: pageHeightPt },
              margin: {
                xPt: state.marginLeft,
                yPt: state.marginTop,
                widthPt: Math.max(0, state.pageWidth - state.marginLeft - state.marginRight),
                heightPt: Math.max(0, pageHeightPt - state.marginTop - state.marginBottom),
              },
              column: request.location.availableBounds,
            },
            floatingTableRegistry: floatRegistry,
            finalPlacementTranslationPt: {
              xPt: request.location.availableBounds.xPt,
              yPt: request.location.cursorPt.yPt,
            },
            reacquirePageDependentBlock: reacquireTableBlock,
          });
          const tableInlineStartPt = request.location.availableBounds.xPt
            + retained.layout.flowBounds.xPt;
          const tableInlineEndPt = tableInlineStartPt + retained.layout.flowBounds.widthPt;
          const remainingTableExtentPt = result.fragment?.advancePt ?? 0;
          let retryAtBlockStartPt = request.location.cursorPt.yPt;
          while (true) {
            const blockers = floatRegistry.entries.filter((entry) => entry.kind === 'table'
              && tableInlineEndPt - entry.exclusionBounds.xPt > FLOAT_OVERLAP_EPS
              && entry.exclusionBounds.xPt + entry.exclusionBounds.widthPt
                - tableInlineStartPt > FLOAT_OVERLAP_EPS
              && retryAtBlockStartPt + remainingTableExtentPt
                - entry.exclusionBounds.yPt > FLOAT_OVERLAP_EPS
              && entry.exclusionBounds.yPt + entry.exclusionBounds.heightPt
                - retryAtBlockStartPt > FLOAT_OVERLAP_EPS);
            if (blockers.length === 0) break;
            retryAtBlockStartPt = Math.max(...blockers.map((entry) => (
              entry.exclusionBounds.yPt + entry.exclusionBounds.heightPt
            )));
          }
          if (retryAtBlockStartPt > request.location.cursorPt.yPt) {
            return Object.freeze({
              layout: retained.layout,
              blockExtentPt: 0,
              nextCursor: request.cursor ?? null,
              retryAtBlockStartPt,
            });
          }
          if (!result.fragment || result.requiresFreshPage) {
            return Object.freeze({
              layout: retained.layout,
              blockExtentPt: 0,
              nextCursor: Object.freeze({ kind: 'table' as const, cursor }),
              requiresFreshFlowRegion: true,
            });
          }
          return Object.freeze({
            layout: result.fragment,
            blockExtentPt: result.fragment.advancePt,
            nextCursor: result.nextCursor
              ? Object.freeze({ kind: 'table' as const, cursor: result.nextCursor })
              : null,
            ...(result.floatingTableRegistryDelta
              ? {
                  flowRegistryDelta: Object.freeze({
                    floats: result.floatingTableRegistryDelta,
                  }),
                }
              : {}),
          });
        },
        layoutStory: acquireStoryLayout,
        layoutNotes(request) {
          const notes: NoteLayout[] = [];
          let cursorYPt = request.container.bounds.yPt;
          let first = request.firstOnPage;
          for (const id of request.referenceIds) {
            const sourceNotes = request.kind === 'footnote' ? footnotesById : endnotesById;
            if (!sourceNotes.has(id)) continue;
            const source: SourceRef = {
              story: request.kind,
              storyInstance: id,
              path: [],
            };
            const separatorHeightPt = first ? FOOTNOTE_SEPARATOR_GAP_PT : 0;
            const storyContainer = {
              ...request.container,
              id: `${request.container.id}:${request.kind}:${id}`,
              bounds: {
                ...request.container.bounds,
                yPt: cursorYPt + separatorHeightPt,
                heightPt: Math.max(
                  0,
                  request.container.bounds.yPt + request.container.bounds.heightPt
                    - cursorYPt - separatorHeightPt,
                ),
              },
            };
            let story: StoryLayout;
            try {
              story = acquireStoryLayout({
                source,
                pageIndex: request.pageIndex,
                section: request.section,
                container: storyContainer,
              });
            } catch (error) {
              if (error instanceof FlowCapacityExceededError
                && error.containerId === storyContainer.id) {
                throw new NoteCapacityExceededError(
                  request.kind,
                  request.pageIndex,
                  request.container.id,
                );
              }
              throw error;
            }
            const separator = first ? Object.freeze([Object.freeze({
              edge: 'top' as const,
              from: Object.freeze({
                xPt: request.container.bounds.xPt,
                yPt: cursorYPt + separatorHeightPt / 2,
              }),
              to: Object.freeze({
                xPt: request.container.bounds.xPt + request.container.bounds.widthPt / 3,
                yPt: cursorYPt + separatorHeightPt / 2,
              }),
              color: '#000000',
              widthPt: 0.5,
              authoredStyle: 'single',
              style: 'solid' as const,
            })]) : Object.freeze([]);
            const advancePt = separatorHeightPt + story.advancePt;
            const flowBounds = Object.freeze({
              xPt: request.container.bounds.xPt,
              yPt: cursorYPt,
              widthPt: request.container.bounds.widthPt,
              heightPt: advancePt,
            });
            const note: NoteLayout = Object.freeze({
              kind: 'note',
              id: `${request.kind}:${id}:page:${request.pageIndex}`,
              source,
              flowDomainId: request.container.id,
              ordinaryFlow: true,
              flowBounds,
              inkBounds: Object.freeze({
                xPt: Math.min(flowBounds.xPt, story.inkBounds.xPt),
                yPt: Math.min(flowBounds.yPt, story.inkBounds.yPt),
                widthPt: Math.max(
                  flowBounds.xPt + flowBounds.widthPt,
                  story.inkBounds.xPt + story.inkBounds.widthPt,
                ) - Math.min(flowBounds.xPt, story.inkBounds.xPt),
                heightPt: Math.max(
                  flowBounds.yPt + flowBounds.heightPt,
                  story.inkBounds.yPt + story.inkBounds.heightPt,
                ) - Math.min(flowBounds.yPt, story.inkBounds.yPt),
              }),
              clipBounds: request.container.bounds,
              advancePt,
              separator,
              story,
            });
            notes.push(note);
            cursorYPt += advancePt;
            first = false;
          }
          return Object.freeze(notes);
        },
        measureFollowingBlock(request) {
          const candidate: RenderState = {
            ...state,
            floats: [...state.floats],
            retainedTablesBySourceIndex: new Map(state.retainedTablesBySourceIndex),
          };
          applyLocationTo(candidate, request.location);
          if (request.input.kind === 'adjacent-table-group') {
            const records = request.input.tables.map((tableInput) => {
              const table = sourceElement(tableInput.source);
              if (table.type !== 'table') throw new Error('Following table source kind mismatch');
              const sourceIndex = tableInput.source.path[0]!;
              computeTablePtLayout(candidate, table, request.availableInlineExtentPt, sourceIndex);
              return retainedTableRecord(candidate, sourceIndex).acquisition;
            });
            const combinedInput = ordinaryAcquisitionInputForAdjacentGroup(
              combineAdjacentTableLayoutInputs(
                request.input.logicalSequenceId,
                records.map((record) => record.input),
              ),
            );
            const layout = layoutRetainedTableInput(combinedInput, {
              container: {
                id: request.location.flowDomainId,
                kind: 'body',
                bounds: request.location.availableBounds,
              },
              cursor: request.location.cursorPt,
              availableBounds: request.location.availableBounds,
            }, services).layout;
            return Object.freeze({
              fullExtentPt: layout.advancePt,
              leadContentExtentPt: layout.rows[0]?.advancePt ?? layout.advancePt,
              fullFootnoteReferenceIds: footnoteIdsInRetainedSlice(layout),
              leadFootnoteReferenceIds: footnoteIdsInRetainedSlice({
                ...layout,
                rows: layout.rows.slice(0, 1),
              }),
            });
          }
          const element = sourceElement(request.input.source);
          if (request.input.kind === 'paragraph') {
            if (element.type !== 'paragraph') throw new Error('Following paragraph source kind mismatch');
            const { layout } = acquireBodyParagraphAtLocation(
              candidate,
              element,
              request.input.source,
              request.location,
              request.availableInlineExtentPt,
              false,
              undefined,
              drawingCollisionRegistry.entries,
            );
            const firstLine = layout.lines[0];
            return Object.freeze({
              fullExtentPt: layout.advancePt,
              // keepNext admits the successor's first content line, including
              // any retained wrap displacement before that line begins.
              leadContentExtentPt: firstLine
                ? firstLine.bounds.yPt + firstLine.advancePt - layout.flowBounds.yPt
                : layout.advancePt,
              fullFootnoteReferenceIds: footnoteIdsInRetainedSlice(layout),
              leadFootnoteReferenceIds: firstLine
                ? footnoteIdsInRetainedLines([firstLine])
                : [],
            });
          }
          if (element.type !== 'table') throw new Error('Following table source kind mismatch');
          const sourceIndex = request.input.source.path[0]!;
          computeTablePtLayout(candidate, element, request.availableInlineExtentPt, sourceIndex);
          const layout = retainedTableRecord(candidate, sourceIndex).acquisition.layout;
          return Object.freeze({
            fullExtentPt: layout.advancePt,
            leadContentExtentPt: layout.rows[0]?.advancePt ?? layout.advancePt,
            fullFootnoteReferenceIds: footnoteIdsInRetainedSlice(layout),
            leadFootnoteReferenceIds: footnoteIdsInRetainedSlice({
              ...layout,
              rows: layout.rows.slice(0, 1),
            }),
          });
        },
        prescanPageAnchors(request: PageAnchorPrescanInput) {
          const geometry = request.location.section.geometry;
          const marginTopPt = bodyMarginInsetPt(geometry.marginTop);
          const marginBottomPt = bodyMarginInsetPt(geometry.marginBottom);
          const frames = Object.freeze({
            page: Object.freeze({
              xPt: 0, yPt: 0,
              widthPt: geometry.pageWidth, heightPt: geometry.pageHeight,
            }),
            margin: Object.freeze({
              xPt: geometry.marginLeft, yPt: marginTopPt,
              widthPt: Math.max(0, geometry.pageWidth - geometry.marginLeft - geometry.marginRight),
              heightPt: Math.max(0, geometry.pageHeight - marginTopPt - marginBottomPt),
            }),
            column: Object.freeze({
              xPt: request.location.availableBounds.xPt, yPt: marginTopPt,
              widthPt: request.availableInlineExtentPt,
              heightPt: Math.max(0, geometry.pageHeight - marginTopPt - marginBottomPt),
            }),
            paragraph: null,
            line: null,
            character: null,
            pageParity: request.location.pageIndex % 2 === 0 ? 'odd' as const : 'even' as const,
          });
          const publicParagraphs = new Set<DocParagraph>();
          const paragraphKey = (source: SourceRef) =>
            `${source.story}:${source.storyInstance}:${source.path.join('.')}`;
          const paragraphIds = new Map<string, number>();
          const paragraphIdFor = (source: SourceRef): number => {
            const key = paragraphKey(source);
            if (!paragraphIds.has(key)) {
              paragraphIds.set(key, floatRegistry.nextParagraphId + paragraphIds.size);
            }
            return paragraphIds.get(key)!;
          };
          const entries = request.anchors.flatMap((anchor): readonly FloatRegistryEntryPt[] => {
            const paragraph = sourceElement(anchor.paragraphSource);
            if (paragraph.type !== 'paragraph') {
              throw new Error('Page-anchor prescan source kind mismatch');
            }
            const acquired = paragraphAcquisitionInput(paragraph, anchor.paragraphSource);
            const hostMatches = acquired.runs.filter((run) =>
              run.type === 'anchorHost' && run.anchorOccurrenceId === anchor.occurrenceId);
            const payloads = acquired.runs
              .map((run, runIndex) => ({ run, runIndex }))
              .filter((candidate): candidate is typeof candidate & {
                run: Extract<typeof candidate.run, { type: 'image' | 'chart' | 'shape' }> & {
                  anchorAcquisitionInput: NonNullable<Extract<typeof candidate.run, {
                    type: 'image' | 'chart' | 'shape';
                  }>['anchorAcquisitionInput']>;
                };
              } => (
                (candidate.run.type === 'image'
                  || candidate.run.type === 'chart'
                  || candidate.run.type === 'shape')
                && candidate.run.anchorAcquisitionInput?.occurrenceId === anchor.occurrenceId
              ))
              .sort((left, right) => (
                (left.run.anchorAcquisitionInput.group?.sourceIndex ?? 0)
                  - (right.run.anchorAcquisitionInput.group?.sourceIndex ?? 0)
                  || left.runIndex - right.runIndex
              ));
            if (hostMatches.length !== 1 || payloads.length === 0) {
              const publicRun = paragraph.runs.find((run, runIndex) =>
                publicAnchorBridge(run, anchor.paragraphSource, runIndex)?.occurrenceId
                  === anchor.occurrenceId);
              if (publicRun) {
                if (
                  (publicRun.type === 'image'
                    || publicRun.type === 'chart'
                    || publicRun.type === 'shape')
                  && publicRun.wrapMode === 'none'
                ) return [];
                const candidate: RenderState = {
                  ...state,
                  floats: [...state.floats],
                  pageAnchorPrescanned: new Set(state.pageAnchorPrescanned),
                };
                applyLocationTo(candidate, request.location);
                const publicEntries = publicParagraphFloatAcquisition(
                  paragraph,
                  anchor.paragraphSource,
                  candidate,
                  new Set([anchor.occurrenceId]),
                  paragraphIdFor(anchor.paragraphSource),
                );
                if (publicEntries.length !== 1) {
                  throw new Error(`Public page-anchor prescan occurrence mismatch: ${anchor.occurrenceId}`);
                }
                publicParagraphs.add(paragraph);
                return publicEntries;
              }
              throw new Error(`Page-anchor prescan occurrence acquisition mismatch: ${anchor.occurrenceId}`);
            }
            const result = resolveAnchorFrame({
              acquisition: payloads[0]!.run.anchorAcquisitionInput,
              frames,
            });
            if (result.status !== 'resolved') {
              throw new Error(`Page-anchor prescan could not resolve occurrence: ${anchor.occurrenceId}`);
            }
            const wrapBounds = result.geometry.wrapBounds;
            if (wrapBounds === null || result.geometry.wrap.kind === 'none') {
              return [];
            }
            const polygon = result.geometry.wrap.polygon?.points ?? Object.freeze([
              Object.freeze({ xPt: wrapBounds.xPt, yPt: wrapBounds.yPt }),
              Object.freeze({ xPt: wrapBounds.xPt + wrapBounds.widthPt, yPt: wrapBounds.yPt }),
              Object.freeze({
                xPt: wrapBounds.xPt + wrapBounds.widthPt,
                yPt: wrapBounds.yPt + wrapBounds.heightPt,
              }),
              Object.freeze({ xPt: wrapBounds.xPt, yPt: wrapBounds.yPt + wrapBounds.heightPt }),
            ]);
            return [Object.freeze({
              kind: 'shape' as const,
              occurrenceId: anchor.occurrenceId,
              paragraphId: paragraphIdFor(anchor.paragraphSource),
              bounds: result.geometry.objectFrame,
              exclusionBounds: wrapBounds,
              wrap: result.geometry.wrap.kind,
              wrapSide: result.geometry.wrap.side,
              wrapDistances: result.geometry.wrap.distances,
              wrapPolygon: Object.freeze([...polygon]),
            })];
          });
          publicParagraphs.forEach((paragraph) => state.pageAnchorPrescanned?.add(paragraph));
          if (entries.length === 0) return null;
          return Object.freeze({
            floats: Object.freeze({
              coordinateSpace: 'logical-page-points' as const,
              // Page-owned wrap exclusions survive same-page column/section
              // cutovers, so their transaction identity belongs to the physical
              // page rather than the active body flow domain.
              flowDomainId: floatRegistry.flowDomainId,
              baseEntries: floatRegistry.entries,
              baseNextParagraphId: floatRegistry.nextParagraphId,
              nextParagraphId: floatRegistry.nextParagraphId + entries.length,
              entries: Object.freeze(entries),
            }),
          });
        },
        measureLineNumberGlyph(text) {
          const previousFont = measureContext.font;
          try {
            const fontSizePt = docDefaultFontSizePt(doc);
            const font = buildFont(false, false, fontSizePt, null, {});
            measureContext.font = font;
            const metrics = measureContext.measureText(text);
            return Object.freeze({
              widthPt: metrics.width,
              ascentPt: metrics.fontBoundingBoxAscent
                ?? metrics.actualBoundingBoxAscent
                ?? fontSizePt * 0.8,
              descentPt: metrics.fontBoundingBoxDescent
                ?? metrics.actualBoundingBoxDescent
                ?? fontSizePt * 0.2,
              font,
            });
          } finally {
            measureContext.font = previousFont;
          }
        },
        resetPageAcquisition(next) {
          state.floats = [];
          state.floatParaSeq = 0;
          state.pageAnchorPrescanned = new Set();
          floatRegistry = Object.freeze({
            coordinateSpace: 'logical-page-points' as const,
            flowDomainId: pageRegistryFlowDomainId(next.pageIndex),
            entries: Object.freeze([]),
            nextParagraphId: 0,
          });
          drawingCollisionRegistry = createDrawingMLCollisionRegistry(
            pageRegistryFlowDomainId(next.pageIndex),
            'logical-page-points',
          );
          applyLocation(next);
        },
        moveAcquisitionCursor: applyLocation,
        flowRegistrySnapshot(): BodyFlowRegistrySnapshotPt {
          return Object.freeze({
            floats: floatRegistry,
            drawingCollisions: drawingCollisionRegistry,
          });
        },
        commitFlowRegistryDelta(delta: BodyFlowRegistryDeltaPt) {
          if (!delta.floats && !delta.drawingCollisions) {
            throw new Error('Body flow registry delta must update at least one registry');
          }
          if (delta.floats) {
            validateFloatingTableRegistryDelta(delta.floats, {
              coordinateSpace: floatRegistry.coordinateSpace,
              flowDomainId: floatRegistry.flowDomainId,
              entries: floatRegistry.entries,
              nextParagraphId: floatRegistry.nextParagraphId,
            });
          }
          if (delta.drawingCollisions) {
            validateDrawingMLCollisionRegistryDelta(
              drawingCollisionRegistry,
              delta.drawingCollisions,
            );
          }
          const nextDrawingCollisionRegistry = delta.drawingCollisions
            ? applyDrawingMLCollisionRegistryDelta(
                drawingCollisionRegistry,
                delta.drawingCollisions,
              )
            : drawingCollisionRegistry;
          const scale = state.scale;
          const retainedFloats = (delta.floats?.entries ?? []).map((entry): FloatRect => {
            const left = entry.wrapDistances?.leftPt
              ?? entry.bounds.xPt - entry.exclusionBounds.xPt;
            const top = entry.wrapDistances?.topPt
              ?? entry.bounds.yPt - entry.exclusionBounds.yPt;
            const right = entry.wrapDistances?.rightPt
              ?? entry.exclusionBounds.xPt + entry.exclusionBounds.widthPt
                - entry.bounds.xPt - entry.bounds.widthPt;
            const bottom = entry.wrapDistances?.bottomPt
              ?? entry.exclusionBounds.yPt + entry.exclusionBounds.heightPt
                - entry.bounds.yPt - entry.bounds.heightPt;
            return {
              kind: entry.kind,
              mode: entry.wrap === 'topAndBottom' ? 'topAndBottom' : 'square',
              ...(entry.kind === 'shape' ? {
                anchorOccurrenceId: entry.occurrenceId,
                acquisitionOccurrenceId: entry.occurrenceId,
              } : {}),
              ...(entry.wrap ? {
                authoredWrap: entry.wrap,
                wrapPolygon: entry.wrapPolygon,
              } : {}),
              imageKey: entry.exclusionId
                ?? (entry.kind === 'table' ? `body:float:${entry.paragraphId}` : ''),
              imageX: entry.bounds.xPt * scale, imageY: entry.bounds.yPt * scale,
              imageW: entry.bounds.widthPt * scale, imageH: entry.bounds.heightPt * scale,
              xLeft: entry.exclusionBounds.xPt * scale,
              xRight: (entry.exclusionBounds.xPt + entry.exclusionBounds.widthPt) * scale,
              yTop: entry.exclusionBounds.yPt * scale,
              yBottom: (entry.exclusionBounds.yPt + entry.exclusionBounds.heightPt) * scale,
              side: entry.wrapSide ?? 'bothSides',
              distLeft: left * scale, distRight: right * scale,
              distTop: top * scale, distBottom: bottom * scale,
              paraId: entry.paragraphId, drawn: true,
            };
          });
          if (delta.floats) {
            state.floats.push(...retainedFloats);
            floatRegistry = Object.freeze({
              ...floatRegistry,
              entries: Object.freeze([...floatRegistry.entries, ...delta.floats.entries]),
              nextParagraphId: delta.floats.nextParagraphId,
            });
            state.floatParaSeq = delta.floats.nextParagraphId;
          }
          drawingCollisionRegistry = nextDrawingCollisionRegistry;
        },
      };
      return Object.freeze(session);
    },
  });
}

function paragraphMeasurementEnvironment(
  state: Pick<
    RenderState,
    | 'pageIndex'
    | 'totalPages'
    | 'displayPageNumber'
    | 'pageNumberFormat'
    | 'currentDateMs'
    | 'noteNumbers'
    | 'noteReferenceNumber'
    | 'verticalCJK'
    | 'verticalAllRotated'
    | 'docEastAsian'
    | 'resolvedLocalFonts'
    | 'layoutServices'
  >,
): ParagraphMeasurementEnvironment {
  return {
    pageIndex: state.pageIndex,
    totalPages: state.totalPages,
    displayPageNumber: state.displayPageNumber,
    pageNumberFormat: state.pageNumberFormat,
    currentDateMs: state.currentDateMs,
    noteNumbers: state.noteNumbers,
    noteReferenceNumber: state.noteReferenceNumber,
    // §17.6.20 btLr (#988 re-adjudication): an all-rotated page lays its lines
    // out with the HORIZONTAL semantics (no 縦中横 grouping — the whole layout
    // rotates wholesale), so the environment's vertical flag is the effective
    // "upright vertical" one. tbRl (no verticalAllRotated) is unchanged.
    verticalCJK: state.verticalCJK && !state.verticalAllRotated,
    documentHasEastAsianText: state.docEastAsian,
    resolvedLocalFonts: state.resolvedLocalFonts,
    layoutServices: state.layoutServices,
  };
}

/** The `LineLayoutEnvironment` for building a paragraph's segments straight
 *  from a paint state. Identity (the state itself — byte-identical) for
 *  horizontal and upright-vertical (tbRl family) states; an all-rotated
 *  (§17.6.20 btLr, #988 re-adjudication) page builds its segments with
 *  `verticalCJK` CLEARED so no 縦中横 grouping (§17.3.2.10) engages — the btLr
 *  page is the horizontal layout rotated wholesale, so its segments are the
 *  horizontal ones (measure == paint: the paginator's measure state never sets
 *  `verticalCJK` either). */
function segmentEnvironmentOf(state: RenderState): RenderState {
  return state.verticalAllRotated ? { ...state, verticalCJK: false } : state;
}

function buildParagraphFragment(
  source: DocParagraph,
  measured: MeasuredParagraph,
  lineStart: number,
  lineEnd: number,
  isFirstSlice: boolean,
  isFinalSlice: boolean,
  trailingExtentPt: number,
  state: RenderState,
  sourceRef: SourceRef,
  flowDomainId = 'body',
  paragraphBorderEdges?: NonNullable<Parameters<typeof paragraphLayoutFromMeasurement>[1]['paragraphBorderEdges']>,
  acquiredContext?: ParagraphLayoutContext,
  sourceRangeStart?: number,
): ParagraphLayout {
  const exclusions = paragraphWrapExclusions(state.floats, flowDomainId);
  const id = `${sourceRef.story}:${sourceRef.storyInstance}:${sourceRef.path.join('.')}`;
  const whole = paragraphLayoutFromMeasurement(
    paragraphAcquisitionInput(source, sourceRef),
    {
      id,
      source: sourceRef,
      flowDomainId,
      ordinaryFlow: true,
      context: acquiredContext ?? resolveStateParagraphLayoutContext(state, source),
      placement: measured.placement,
      measurer: { context: state.ctx, fontFamilyClasses: state.fontFamilyClasses },
      environment: paragraphMeasurementEnvironment(state),
      exclusions,
      containerShading: state.containerShading,
      ...(paragraphBorderEdges ? { paragraphBorderEdges } : {}),
      trailingExtentPt,
      continuesFromPrevious: !isFirstSlice,
      ...(sourceRangeStart === undefined ? {} : { sourceRangeStart }),
      anchorFrames: paragraphAnchorReferenceFrames(state),
    },
    measured,
  );
  if (lineStart === 0 && lineEnd === whole.lines.length && isFirstSlice && isFinalSlice) {
    return whole;
  }
  return sliceParagraphLayout(whole, {
    lineStart,
    lineEnd,
    continuesFromPrevious: !isFirstSlice,
    continuesOnNext: !isFinalSlice,
  });
}

function retainedTableRecord(state: RenderState, sourceIndex: number): RetainedTableRecord {
  const record = state.retainedTablesBySourceIndex?.get(sourceIndex);
  if (!record) throw new Error('Table placement requires retained table acquisition');
  return record;
}

function measureBodyParagraphAtCursor(
  state: RenderState,
  para: DocParagraph,
  contentWPt: number,
  suppressSpaceBefore: boolean,
  paraXPt: number,
  continuation?: Readonly<{
    boundary: LineBoundary;
    uniformRubyAdvancePt?: number;
  }>,
): MeasuredParagraph {
  const paragraphContext = resolveBodyParagraphLayoutContext(state, para);
  return measureParagraph(
    para,
    paragraphContext,
    {
      startYPt: state.y,
      paragraphXPt: paraXPt,
      availableWidthPt: contentWPt,
      maximumYPt: state.pageH,
      suppressSpaceBefore,
      wrap: state.floats.length > 0
        ? createFloatWrapOracle(state.floats, {
            xLeftPt: 0,
            xRightPt: state.pageWidth,
            readingDirection: paragraphContext.baseRtl ? 'rtl' : 'ltr',
          })
        : undefined,
    },
    {
      context: state.ctx,
      fontFamilyClasses: state.fontFamilyClasses,
    },
    paragraphMeasurementEnvironment(state),
    continuation,
  );
}

/** The estimated flow height of an already-measured body paragraph: its content span
 *  from the measurement's recorded start, plus trailing space (spaceAfter or the
 *  §17.3.1.7 bottom-border extent, unless the next in-flow paragraph shares the border
 *  box). `measured.placement.startYPt` equals the `state.y` the measurement was taken
 *  at, so this reproduces the original `contentEndYPt − startYPt + …` formula. */
function paragraphHeightFromMeasured(
  measured: MeasuredParagraph,
  para: DocParagraph,
  nextSharesBottomBorder: boolean,
): number {
  const bottomExtent = nextSharesBottomBorder ? 0 : bottomBorderExtentPt(para.borders);
  return measured.contentEndYPt - measured.placement.startYPt
    + Math.max(measured.requestedSpaceAfterPt, bottomExtent);
}

function estimateParagraphHeight(
  state: RenderState,
  para: DocParagraph,
  contentWPt: number,
  suppressSpaceBefore = false,
  paraXPt = 0,
  /** §17.3.1.7: the next in-flow paragraph shares this paragraph's border box, so
   *  its bottom edge is suppressed (the box continues) and reserves no extent. */
  nextSharesBottomBorder = false,
): number {
  // ECMA-376 §17.3.1.29 + §17.3.2.41: a fully-hidden paragraph (inkless +
  // vanished mark) collapses to zero height, so every look-ahead estimate
  // (keepNext's estimateNextBlockHeight, the inline-image-cluster scan) that
  // folds one in stays in lockstep with the paginator's whole-skip above.
  if (isFullyHiddenParagraph(para)) return 0;
  return paragraphHeightFromMeasured(
    measureBodyParagraphAtCursor(state, para, contentWPt, suppressSpaceBefore, paraXPt),
    para,
    nextSharesBottomBorder,
  );
}

/** Snap a paragraph's uniform line height up to an integer multiple of the
 *  docGrid pitch. Mirrors Word's docGrid handling for ruby paragraphs:
 *  the grid pitch widens to accommodate the tallest required line, and
 *  every line in the paragraph then uses that widened pitch. */
function snapParaLineToGrid(h: number, grid: DocGridCtx | undefined, scale: number): number {
  if (!isGridLineRule(grid)) return h;
  const pitchPx = grid!.linePitchPt! * scale;
  if (pitchPx <= 0) return h;
  if (h <= pitchPx) return pitchPx;
  return Math.ceil(h / pitchPx) * pitchPx;
}

/** Return true when any text run in the paragraph carries a `ruby` annotation.
 *  Used to apply paragraph-wide line-height snapping to docGrid pitch — Word
 *  renders the entire ruby paragraph with consistent line spacing so that
 *  ruby-bearing and ruby-free lines line up on the same baseline grid. */
/** The docGrid that governs a paragraph's line heights. ECMA-376 §17.3.1.32:
 *  a paragraph with `w:snapToGrid` explicitly off ignores the section grid, so
 *  its lines use natural font metrics / the spacing multiplier directly. */
function gridForParagraphContext(
  state: Pick<RenderState, 'docGrid'>,
  context: ParagraphLayoutContext,
): DocGridCtx {
  return {
    type: state.docGrid.type,
    linePitchPt: context.lineGrid.active ? context.lineGrid.pitchPt : null,
    charSpacePt:
      context.characterGrid.active ? context.characterGrid.deltaPt : null,
  };
}

function paraGrid(para: DocParagraph, state: RenderState): DocGridCtx {
  return gridForParagraphContext(
    state,
    resolveStateParagraphLayoutContext(state, para),
  );
}

function computeTableRowHeights(
  state: RenderState,
  table: DocTable,
  contentWPt: number,
  sourceIndex?: number,
): number[] {
  return computeTablePtLayout(state, table, contentWPt, sourceIndex).rowHeightsPt;
}

/** The paginator's scale-1 table layout: the per-grid-column widths (pt) and the
 *  per-row heights (pt), both resolved through the SAME functions the paint pass
 *  uses ({@link resolveColumnWidths} + {@link resolveTableRowHeights} with the
 *  unified {@link measureCellContentHeightPx} at scale 1). Returned together so
 *  the paginator can stamp both onto the table element (B2 table stage 1b) for the
 *  paint pass to reuse — one column resolution feeds both the stamp and the row
 *  heights, so the min-content scan runs once. */
function computeTablePtLayout(
  state: RenderState,
  table: DocTable,
  contentWPt: number,
  sourceIndex?: number,
): { colWidthsPt: number[]; rowContentHeightsPt: number[]; rowHeightsPt: number[] } {
  const prior = sourceIndex === undefined
    ? undefined
    : state.retainedTablesBySourceIndex?.get(sourceIndex);
  const colWidthsPt = resolveColumnWidths(table, contentWPt, state);
  const dependencies = state.retainedTableAcquisition;
  if (dependencies && sourceIndex !== undefined) {
    const acquired = acquireRetainedTable(
      table,
      colWidthsPt,
      contentWPt,
      state,
      [sourceIndex],
      dependencies,
    );
    // Split rows are page-local acquisitions, but an unchanged inline extent
    // retains one authoritative track vector for the table's full occurrence.
    const retained = prior?.contentWidthPt === contentWPt
      ? Object.freeze({
          ...acquired,
          layout: Object.freeze({
            ...acquired.layout,
            columnWidthsPt: prior.acquisition.layout.columnWidthsPt,
          }),
        })
      : acquired;
    state.retainedTablesBySourceIndex?.set(sourceIndex, Object.freeze({
      sourceIndex,
      acquisition: retained,
      contentWidthPt: contentWPt,
      anchorYPt: state.y,
    }));
    const rowHeightsPt = retained.layout.rows.map((row) => row.advancePt);
    return { colWidthsPt, rowContentHeightsPt: rowHeightsPt, rowHeightsPt };
  }
  const rowContentHeightsPt = resolveTableRowContentHeights(table, colWidthsPt, 1, (cell, cellW) =>
    measureCellContentHeightPx(cell, table, cellW, 1, state),
  );
  const rowHeightsPt = applyTableRowBoundaryFootprints(table, rowContentHeightsPt, 1);
  return { colWidthsPt, rowContentHeightsPt, rowHeightsPt };
}

function estimateTableHeight(
  state: RenderState,
  table: DocTable,
  contentWPt: number,
  sourceIndex?: number,
): number {
  return computeTableRowHeights(state, table, contentWPt, sourceIndex).reduce((s, x) => s + x, 0);
}

/**
 * Acquire the parser/style facts and intrinsic content constraints required by
 * ECMA-376 §17.18.87, then resolve the shared table grid. `tblGrid` is the
 * initial grid, not an oracle containing an application's previous result;
 * authored `tblW`, `tcW`, `wBefore`, and `wAfter` remain active constraints.
 * Exported for the table-layout integration tests.
 */
export function resolveColumnWidths(table: DocTable, contentWPt: number, state: RenderState): number[] {
  const format = tableFormatInput(table);
  const marginsByCell = new WeakMap<object, Readonly<{ left: number; right: number }>>();
  table.rows.forEach((row, rowIndex) => row.cells.forEach((cell, cellIndex) => {
    const acquired = format.rows[rowIndex]?.cells[cellIndex]?.marginsPt;
    marginsByCell.set(cell, acquired ?? effCellMargins(cell, table));
  }));
  return [...resolveTableColumnWidths(tableColumnLayoutInput(
    table,
    contentWPt,
    (cell) => {
      const margins = marginsByCell.get(cell as object) ?? effCellMargins(cell as DocTableCell, table);
      return measureTableCellIntrinsicWidths(cell, margins, {
        paragraph: (paragraph) => {
          // This exported low-level seam historically accepts a reduced state.
          // Production uses normalized contexts; the fallback preserves only
          // the stable hand-built test/input contract at this renderer adapter.
          const baseContext: ParagraphLayoutContext = state.layoutSettings && state.sectionLayout
            ? resolveParagraphLayoutContext(
                state.layoutSettings,
                state.sectionLayout,
                state.storyContext ?? BODY_STORY_CONTEXT,
                paragraph,
              )
            : {
                lineGrid: { active: false, pitchPt: null },
                characterGrid: { active: false, deltaPt: 0 },
                physicalIndentLeftPt: paragraph.bidi ? paragraph.indentRight : paragraph.indentLeft,
                physicalIndentRightPt: paragraph.bidi ? paragraph.indentLeft : paragraph.indentRight,
                firstIndentPt: paragraph.indentFirst,
                lineSpacing: paragraph.lineSpacing,
                spaceBeforePt: paragraph.spaceBefore,
                spaceAfterPt: paragraph.spaceAfter,
                baseRtl: paragraph.bidi === true,
                isJustified: jcIsFullyJustified(paragraph.alignment),
                stretchLastLine: jcStretchesLastLine(paragraph.alignment),
                tabStops: [...paragraph.tabStops],
                hasRuby: paragraph.runs.some(
                  (run) => run.type === 'text' && Boolean((run as DocxTextRun).ruby),
                ),
                hasEastAsianText: paragraph.runs.some(
                  (run) => run.type === 'text' && EAST_ASIAN_RE.test((run as DocxTextRun).text),
                ),
                kinsoku: state.kinsoku ?? DEFAULT_KINSOKU_RULES,
                defaultTabPt: state.defaultTabPt ?? DEFAULT_TAB_PT,
              };
          const markerInput = paragraph.numbering
            ? numberingMarkerShapeInput(paragraph.numbering, getDefaultFontSize(paragraph))
            : undefined;
          const context = applyNumberingBodyOffset(baseContext, {
            numbering: paragraph.numbering,
            ...(markerInput ? { markerInput } : {}),
            authoredFirstIndentPt: paragraph.indentFirst,
            tabStops: paragraph.tabStops,
            defaultTabPt: state.defaultTabPt,
            service: state.layoutServices?.text,
            clusterGeometry: false,
          });
          const numbering = context.numberingMarkerGeometry
            ?? (paragraph.numbering && markerInput && state.layoutServices?.text
              ? resolveNumberingMarkerGeometry(paragraph.numbering, markerInput, {
                  authoredFirstIndentPt: paragraph.indentFirst,
                  physicalIndentLeftPt: context.physicalIndentLeftPt,
                  tabStops: paragraph.tabStops,
                  defaultTabPt: state.defaultTabPt,
                }, state.layoutServices.text, false)
              : undefined);
          return measureParagraphIntrinsicWidths(
            paragraph,
            context,
            contentWPt,
            { context: state.ctx, fontFamilyClasses: state.fontFamilyClasses },
            paragraphMeasurementEnvironment(state),
            numbering,
          );
        },
        nestedTable: (nested) => {
          const widthPt = resolveColumnWidths(nested, contentWPt, state)
            .reduce((sum, width) => sum + width, 0);
          return { minWidthPt: widthPt, maxWidthPt: widthPt };
        },
      });
    },
    tableParticipatesInOrdinaryFlow(table)
      ? contentWPt
      : Math.max(contentWPt, state.pageWidth),
  ))];
}

/** A page break before row `ri` is unsafe when `ri` continues a vertical merge
 *  started above (ECMA-376 §17.4.85): splitting there would orphan the merged
 *  cell's continuation. Such a row carries at least one `vMerge=false` cell.
 *
 *  When an over-tall vMerge span is broken at an interior boundary (see the
 *  {@link splitTableAcrossPages} relaxation), the continuation slice re-opens the
 *  merged cell via {@link reopenMergedCellsInRow} so this rule is re-satisfied for
 *  the slice as its own table. */
function tableBreakAllowedBefore(table: DocTable, ri: number): boolean {
  if (ri <= 0) return true;
  return !table.rows[ri].cells.some((c) => c.vMerge === false);
}

/** The cell whose gridSpan covers logical column `targetCi` in `row`, or `null`.
 *  Pure grid walk (gridSpan-aware), mirroring {@link findMergeEndRow}'s column
 *  scan (ECMA-376 §17.4.85). */
function cellAtGridColumn(
  row: DocTableRow,
  targetCi: number,
  columnCount: number,
): DocTableCell | null {
  let ci = rowGridBefore(row, columnCount);
  for (const cell of row.cells) {
    if (targetCi >= ci && targetCi < ci + cell.colSpan) return cell;
    ci += cell.colSpan;
  }
  return null;
}

/** ECMA-376 §17.4.85 — re-open a vMerge span that crosses INTO a continuation
 *  slice. When an over-tall span is broken at an interior row boundary, the
 *  slice's first body row (`rows[start]`) inherits `vMerge=continue` cells whose
 *  `restart` sits on an earlier page. {@link drawTableRows} skips a bare continue
 *  cell ("drawn by its restart partner"), so without this the re-opened cell box
 *  would not be painted at all. For each such continue cell, walk UP `rows` to its
 *  owning restart and:
 *   - if that restart is a REPEATED header row already prepended to this slice
 *     (`restartRi < headerCount`, and headers repeat), leave the continue cell as
 *     is — the prepended header restart already spans the body rows, so promoting
 *     here would draw a SECOND box (review finding, §17.4.78);
 *   - otherwise promote it to `restart`, cloning the OWNING RESTART cell's
 *     presentation (background / borders / vAlign) so the continuation box matches
 *     Word — which paints the whole merged span from the restart cell — rather than
 *     the continue cell's own (usually empty) properties. Content is dropped: the
 *     merged content stayed with the restart row on the first piece, so the re-
 *     opened box is empty (no duplication). The grid footprint (`colSpan`) is kept
 *     from the continue cell so the row's column math is unchanged.
 *  Runtime-only clone: the parsed rows/cells are never mutated. */
/**
 * ECMA-376 §17.3.1.9 `<w:contextualSpacing>` — Word-adjudicated PER-SIDE
 * semantics (issue #1015, fixture sample-57 ground truth; identical in body,
 * table cell, and text box).
 *
 * The §17.3.1.33 collapsed inter-paragraph gap decomposes as
 *   gap = prevContrib + currContrib
 *   prevContrib = prev.spaceAfter                          (the collapse base)
 *   currContrib = max(curr.spaceBefore − prev.spaceAfter, 0)   (the excess)
 * summing to max(after, before). A paragraph whose toggle is set AND whose
 * neighbour shares its paragraph style drops ITS OWN contribution only:
 *   - prev toggles → gap = max(before − after, 0)  — matches the spec's worked
 *     example (after=10pt, before=12pt → 2pt);
 *   - curr toggles → gap = after — Word renders prev's spaceAfter intact (the
 *     spec-literal "subtract this paragraph's before from the net" would give
 *     0; Word measured 10pt on sample-57 case 2);
 *   - both toggle  → gap = 0 (each side's contribution dropped; the
 *     decomposition is non-negative so no explicit floor is needed).
 *
 * Every gap site uses the effBefore/overlap form
 *   gap = prevAfter + (suppressBefore ? 0 : spaceBefore) − overlap
 * so this helper returns that pair. Kept structural (not `DocParagraph`) so the
 * SAME rule drives the body paragraph path, the cell paths, and the text-box
 * path ({@link ShapeText}), which carry the identical
 * `contextualSpacing`/`styleId` pair.
 */
function contextualSpacingAdjust(
  prev: { contextualSpacing?: boolean; styleId?: string | null } | null,
  curr: { contextualSpacing?: boolean; styleId?: string | null },
  prevAfter: number,
  spaceBefore: number,
): { suppressBefore: boolean; overlap: number } {
  return paragraphGapAdjustment(prev, curr, prevAfter, spaceBefore);
}

/**
 * Whether a paragraph places NO inline content — no text, image, shape, math, or
 * break run. It still produces one paragraph-mark line box (§17.3.1.29), but
 * carries no glyphs.
 */
function isInklessParagraph(p: DocParagraph): boolean {
  return !(p.runs ?? []).some((r) => {
    const run = r as { type?: string; text?: string };
    if (run.type === 'text') return (run.text ?? '').length > 0;
    return true; // image / shape / math / break runs are visible content
  });
}

/**
 * ECMA-376 §17.3.1.29 + §17.3.2.41 — a paragraph with no visible inline content
 * whose paragraph MARK is vanished (hidden text). In the normal/print view
 * (settings hidden-text off — the view a Word PDF export renders) it is not
 * displayed at all, so it collapses to zero height: no mark line box, no
 * paragraph spacing, nothing painted. This is the paragraph-mark analogue of the
 * parser stripping hidden runs (`fmt.vanish` in parser.rs): a run of such empty
 * vanished paragraphs must not reserve vertical space (sample-28, issue #868 —
 * seven of them otherwise forced one extra page). A paragraph with VISIBLE
 * content and a vanished mark is NOT collapsed (it is not inkless): its content
 * still draws, only the pilcrow is hidden.
 */
function isFullyHiddenParagraph(p: DocParagraph): boolean {
  return p.markVanish === true && isInklessParagraph(p);
}

function isAnchorOnlyParagraph(p: DocParagraph): boolean {
  let hasAnchor = false;
  for (const r of p.runs ?? []) {
    if (r.type === 'text' && ((r as DocxTextRun).text ?? '').length === 0) continue;
    if (r.type === 'shape') {
      hasAnchor = true;
      continue;
    }
    if (r.type === 'image' && !!(r as unknown as ImageRun).anchor) {
      hasAnchor = true;
      continue;
    }
    if (r.type === 'chart' && !!(r as unknown as ChartRun).anchor) {
      hasAnchor = true;
      continue;
    }
    return false;
  }
  return hasAnchor;
}

function hasInlineImage(p: DocParagraph): boolean {
  return (p.runs ?? []).some((r) => r.type === 'image' && !(r as unknown as ImageRun).anchor);
}

/**
 * Sum the heights of a cell's content elements with paragraph spacing collapsed
 * the same way `renderCellContent` paints them, so a cell measured for row
 * sizing equals the height it actually paints. Two collapse rules apply
 * (mirroring the paint pass):
 *
 *   - ECMA-376 §17.3.1.9 `<w:contextualSpacing>`: a same-style toggling
 *     paragraph drops its OWN contribution to the inter-paragraph gap
 *     ({@link contextualSpacingAdjust} — per-side, Word-adjudicated).
 *   - Adjacent-paragraph spacing OVERLAP: the gap between two paragraphs is
 *     `max(prevSpaceAfter, currSpaceBefore)`, not their sum. We subtract the
 *     overlap `min(prevSpaceAfter, effBefore)` so a 12pt space-after followed
 *     by a 12pt space-before contributes 12pt of gap, not 24pt.
 *
 * A nested table (CellElement other than paragraph) resets the
 * prev-paragraph context — the next paragraph after a table spaces from a
 * fresh baseline, exactly as `renderCellContent` does.
 *
 * `perElementHeight(elem)` returns each element's full measured height: for a
 * paragraph it must include its full `spaceBefore` (no contextual or overlap
 * adjustment) so this helper can subtract the collapse correctly; for a nested
 * table it is the table's own total height. `spaceScale` converts spec spacing
 * (pt) into the same units as `perElementHeight` returns (1 for pt; the device
 * scale for px); the subtracted collapse therefore lands in matching units.
 *
 * Exported for unit tests (table-spacing-collapse.test).
 */
export function sumCellContentHeight(
  content: CellElement[],
  perElementHeight: (el: CellElement) => number,
  spaceScale: number,
): number {
  let h = 0;
  let prevPara: DocParagraph | null = null;
  let prevSpaceAfter = 0;
  for (const ce of content) {
    if (ce.type === 'paragraph') {
      const para = ce as unknown as DocParagraph;
      // §17.3.1.9 per-side contextualSpacing (contextualSpacingAdjust) over the
      // §17.3.1.33 max-collapse — a same-style toggle drops the toggling
      // paragraph's own contribution to the gap.
      const adjust = contextualSpacingAdjust(prevPara, para, prevSpaceAfter, para.spaceBefore);
      h += perElementHeight(ce)
        - (adjust.suppressBefore ? para.spaceBefore : 0) * spaceScale
        - adjust.overlap * spaceScale;
      prevPara = para;
      prevSpaceAfter = para.spaceAfter;
    } else {
      h += perElementHeight(ce);
      prevPara = null;
      prevSpaceAfter = 0;
    }
  }
  return h;
}

function renderParaList(paras: DocParagraph[], state: RenderState): void {
  let prevPara: DocParagraph | null = null;
  let prevSpaceAfter = 0;
  for (let i = 0; i < paras.length; i++) {
    const para = paras[i];
    // §17.3.1.9 per-side contextualSpacing over the §17.3.1.33 max-collapse —
    // the same contextualSpacingAdjust pair every other flow applies.
    const adjust = contextualSpacingAdjust(prevPara, para, prevSpaceAfter, para.spaceBefore);
    const suppress = adjust.suppressBefore;
    state.y -= adjust.overlap * state.scale;
    // ECMA-376 §17.3.1.7 paragraph-border merge. This list is a single flow (a
    // note or a table cell), so adjacency is just consecutive list members; a
    // frame paragraph (§17.3.1.11) is out of flow and breaks the run. `framePr`
    // siblings are filtered by parasShareBorderBox, so compare with the literal
    // neighbors here.
    const prevSibling = (paras[i - 1] ?? null) as DocParagraph | null;
    const nextSibling = (paras[i + 1] ?? null) as DocParagraph | null;
    const borderMerge: ParaBorderMerge | undefined = hasAnyBorderEdge(para.borders)
      ? {
          suppressTop: parasShareBorderBox(prevSibling, para),
          suppressBottom: parasShareBorderBox(para, nextSibling),
        }
      : undefined;
    renderParagraph(para, state, suppress, undefined, false, borderMerge);
    prevPara = para;
    prevSpaceAfter = para.spaceAfter;
  }
}

// ===== Paragraph rendering =====

/**
 * Map an ST_Jc (math) value to a physical alignment edge for a display
 * equation. ECMA-376 §22.1.2.88: `left`→left, `right`→right, `center` and
 * `centerGroup`→center (for a single equation centerGroup is identical to
 * center; the group-block distinction is out of scope — see spec YAGNI). The
 * math jc is an absolute position within the block, not a logical start/end, so
 * it is NOT flipped by the paragraph base direction.
 */
function mathJcToEdge(jc: string): AlignEdge {
  switch (jc) {
    case 'left':
      return 'left';
    case 'right':
      return 'right';
    case 'center':
    case 'centerGroup':
    default:
      return 'center';
  }
}

/**
 * Render a paragraph that produces NO inline lines — either literally empty
 * (no segments) or anchor-only (its only content is wrap floats, drawn
 * separately). Per ECMA-376 §17.3.1.29 such a paragraph still emits ONE
 * paragraph-mark line box; this advances `state.y` past it, draws its shading /
 * borders, and lays its wrapNone anchor images at the (possibly float-flowed)
 * paragraph base.
 *
 * Shared by renderParagraph's `segments.length === 0` and `lines.length === 0`
 * branches (previously duplicated verbatim). The anchor-only branch's
 * slice-boundary guards (spaceAfter only on the final slice, anchor images only
 * on the first) are parameterized via `markCtx.totalLines` / `lineSlice`; for
 * the literally-empty branch `lineSlice` is always undefined (empty paragraphs
 * are never sliced), so those guards reduce to the unconditional behavior it had.
 */
function renderEmptyMarkParagraph(
  para: DocParagraph,
  state: RenderState,
  markCtx: {
    grid: DocGridCtx;
    paraHasRuby: boolean;
    contentX: number;
    indLeft: number;
    paraW: number;
    borderX: number;
    borderW: number;
    textAreaTopY: number;
    paragraphStartY: number;
    /** Flowed top of the mark line (output of resolveEmptyMarkTop). */
    markTop: number;
    /** Total laid-out line count (0 here); used by the slice guards. */
    totalLines: number;
    lineSlice?: { start: number; end: number; continues?: boolean };
    /** §17.3.1.7 paragraph-border merge (suppress top/bottom edges). */
    borderMerge?: ParaBorderMerge;
  },
): void {
  const { ctx, scale, dryRun } = state;
  const { grid, paraHasRuby, contentX, indLeft, paraW, borderX, borderW, textAreaTopY,
    paragraphStartY, markTop, totalLines, lineSlice, borderMerge } = markCtx;
  // Displacement applied by the float-flow (0 when the mark fits where it is).
  const flowShift = Math.max(0, markTop - textAreaTopY);
  if (markTop > state.y) state.y = markTop;
  const markRectTop = state.y;
  const emptyH = paragraphMarkLineHeight(
    para, scale, grid, paraHasRuby, state.docEastAsian, ctx, state.fontFamilyClasses,
    para.lineSpacing, state.resolvedLocalFonts, state.layoutServices?.text,
    paragraphMarkShapeInput(para),
  );
  if (para.shading && !dryRun) {
    ctx.fillStyle = `#${para.shading}`;
    const sb = paraShadingRect(borderX, markRectTop, borderW, emptyH, para.borders, borderMerge, scale);
    ctx.fillRect(sb.x, sb.y, sb.w, sb.h);
  }
  state.y += emptyH;
  if (para.borders && !dryRun) {
    drawParaBorders(ctx, borderX, markRectTop, borderW, emptyH, para.borders, scale, state.dpr, borderMerge);
  }
  // Only the slice covering the FINAL line emits spaceAfter. With no inline
  // lines there is a single slice, so this is the whole paragraph. §17.3.1.7: a
  // drawn bottom border extends `space + width/2` below the mark box; reserve the
  // amount it pokes past spaceAfter (MAX) so the next paragraph clears it — mirrors
  // estimateParagraphHeight, keeping paint and pagination in lockstep.
  const isFinalSlice = !lineSlice || lineSlice.end >= totalLines;
  if (isFinalSlice) {
    state.y += Math.max(para.spaceAfter, bottomBorderExtentPt(para.borders, borderMerge)) * scale;
  }
  // wrapNone anchor images anchor relative to the paragraph (ayFromPara); when
  // the mark line flowed below a float band the paragraph (and its wrapNone
  // image) drops by the same amount, so shift the anchor base by flowShift while
  // keeping the un-flowed base (paragraphStartY) otherwise unchanged. Wrap
  // shapes are themselves the float band, so they stay anchored to the original
  // paragraph top (§20.4.3.5) instead of following the paragraph mark's flow.
  // Only the first slice draws them (a continuation slice already did on its page).
  if (!lineSlice || (lineSlice.start === 0 && !lineSlice.continues)) {
    renderAnchorImages(para, state, paragraphStartY + flowShift, 'front', paragraphStartY);
  }
}

// ===== Text frames & drop caps (ECMA-376 §17.3.1.11) =====

/**
 * One line height (px) of the anchor (following non-frame) paragraph, used to
 * size a drop cap by `lines` (§17.3.1.11). The drop cap height equals
 * `lines` × this. Scans `elements` after the frame element for the first
 * non-frame paragraph; falls back to the frame paragraph's own single-line
 * height when none follows (a degenerate trailing frame).
 */
function frameAnchorLineHeightPx(
  elements: BodyElement[],
  frameEl: BodyElement,
  state: RenderState,
): number {
  const start = elements.indexOf(frameEl);
  for (let j = start + 1; j < elements.length; j++) {
    const e = elements[j];
    if (e.type !== 'paragraph') continue;
    const p = e as unknown as DocParagraph;
    if (p.framePr) continue; // adjacent frame paragraphs are part of the frame
    return paragraphMarkLineHeight(
      p,
      state.scale,
      paraGrid(p, state),
      resolveBodyParagraphLayoutContext(state, p).hasRuby,
      state.docEastAsian,
      state.ctx,
      state.fontFamilyClasses,
      p.lineSpacing,
      state.resolvedLocalFonts,
      state.layoutServices?.text,
      paragraphMarkShapeInput(p),
    );
  }
  const fp = frameEl as unknown as DocParagraph;
  return paragraphMarkLineHeight(
    fp,
    state.scale,
    paraGrid(fp, state),
    resolveBodyParagraphLayoutContext(state, fp).hasRuby,
    state.docEastAsian,
    state.ctx,
    state.fontFamilyClasses,
    fp.lineSpacing,
    state.resolvedLocalFonts,
    state.layoutServices?.text,
    paragraphMarkShapeInput(fp),
  );
}

/** Resolve a prepared body frame group and attach its retained member layouts. */
function resolveFrameBox(
  para: DocParagraph,
  state: RenderState,
  anchorLineHpx: number,
  onAcquired?: (acquired: ReturnType<typeof acquireRetainedFrameGroup>) => void,
): FrameBox {
  const group = bodyFrameGroupFor(para);
  if (group && state.dryRun) {
    const measurer = { context: state.ctx, fontFamilyClasses: state.fontFamilyClasses };
    const environment = paragraphMeasurementEnvironment(state);
    const borderEdges = group.members.map(bodyParagraphBorderEdgesFor);
    const scale = state.scale;
    const horizontalBand = frameXContainer(group.framePr.hAnchor, state);
    const pointPlacement = {
      contentXPt: state.contentX / scale,
      contentWidthPt: state.contentW / scale,
      pageHeightPt: state.pageH / scale,
      yPt: state.y / scale,
      anchorLineHeightPt: anchorLineHpx / scale,
    };
    const acquired = acquireRetainedFrameGroup(group, {
      contexts: group.members.map((paragraph) =>
        resolveBodyParagraphLayoutContext(state, paragraph)),
      inputs: group.members.map((paragraph, index) =>
        paragraphAcquisitionInput(paragraph, {
          story: 'body', storyInstance: 'body', path: [group.sourceIndices[index]!],
        })),
      borderEdges,
      borderExtentsPt: group.members.map((paragraph, index) =>
        borderEdges[index]?.bottom === 'none' ? 0 : bottomBorderExtentPt(paragraph.borders)),
      measurer,
      environment,
      containerShading: state.containerShading,
      maximumWidthPt: Math.max(0, horizontalBand.right - horizontalBand.left) / scale,
      acquisitionSession: state,
      placementSignature: [
        pointPlacement.contentXPt,
        pointPlacement.contentWidthPt,
        pointPlacement.pageHeightPt,
        pointPlacement.yPt,
        pointPlacement.anchorLineHeightPt,
        state.pageWidth,
        state.marginLeft,
        state.marginRight,
        state.marginTop,
        state.marginBottom,
      ].join('|'),
      place: (contentWidthPt, contentHeightPt) => {
        const box = computeFrameBox(
          group.framePr,
          state,
          pointPlacement.yPt * scale,
          contentWidthPt * scale,
          contentHeightPt * scale,
          pointPlacement.anchorLineHeightPt * scale,
        );
        return Object.freeze({
          bounds: Object.freeze({
            xPt: box.x / scale,
            yPt: box.y / scale,
            widthPt: box.w / scale,
            heightPt: box.h / scale,
          }),
          exclusionBounds: Object.freeze({
            xPt: box.exLeft / scale,
            yPt: box.exTop / scale,
            widthPt: (box.exRight - box.exLeft) / scale,
            heightPt: (box.exBottom - box.exTop) / scale,
          }),
        });
      },
      anchorFrames: paragraphAnchorReferenceFrames(state),
    });
    onAcquired?.(acquired);
    const box: FrameBox = {
      x: acquired.box.bounds.xPt * scale,
      y: acquired.box.bounds.yPt * scale,
      w: acquired.box.bounds.widthPt * scale,
      h: acquired.box.bounds.heightPt * scale,
      exLeft: acquired.box.exclusionBounds.xPt * scale,
      exTop: acquired.box.exclusionBounds.yPt * scale,
      exRight: (acquired.box.exclusionBounds.xPt + acquired.box.exclusionBounds.widthPt) * scale,
      exBottom: (acquired.box.exclusionBounds.yPt + acquired.box.exclusionBounds.heightPt) * scale,
      registerExclusion: true,
      exclusionId: acquired.box.exclusionId,
    };
    return para === group.owner
      ? box
      : { ...box, registerExclusion: false };
  }
  // B1 compatibility scope: only the header/footer story adapter reaches this
  // branch. Body pagination prepares a §17.3.1.11 group for every frame and is
  // structurally prevented from calling the legacy painter by the boundary gate.
  const fp = para.framePr!;
  const { scale } = state;
  const paraTop = state.y;
  const grid = paraGrid(para, state);
  const paragraphContext = resolveBodyParagraphLayoutContext(state, para);
  const paraHasRuby = paragraphContext.hasRuby;
  const segments = buildSegments(para.runs, segmentEnvironmentOf(state));
  const measureW = 100000;
  const lines = segments.length === 0 ? [] : layoutLines(
    state.ctx,
    segments,
    measureW,
    0,
    scale,
    para.tabStops,
    undefined,
    state.fontFamilyClasses,
    0,
    state.kinsoku,
    gridCharDeltaPx(grid, scale),
    state.defaultTabPt,
    measureW,
    paragraphContext.baseRtl,
    paragraphContext.isJustified,
    paragraphContext.stretchLastLine,
  );
  const contentW = lines.length === 0 ? 0 : Math.max(...lines.map((line) =>
    line.segments.reduce((sum, segment) => sum + segment.measuredWidth, 0)));
  const contentH = lines.reduce((sum, line) => sum + lineBoxHeight(
    para.lineSpacing,
    line.ascent,
    line.descent,
    scale,
    grid,
    paraHasRuby,
    line.intendedSingle,
    paraHasRuby ? paragraphContext.hasEastAsianText : (line.eastAsian ?? false),
    line.gridCountSingle,
  ), 0);
  return computeFrameBox(fp, state, paraTop, contentW, contentH, anchorLineHpx);
}

// ===== Floating tables (ECMA-376 §17.4.57 w:tblpPr / §17.4.56 w:tblOverlap) =====
// Placement math (computeFloatTableBox / registerTableFloat / floatTableWrapSide)
// lives in float-table-geometry.ts; the float-table render path below consumes it.

/**
 * Render a paragraph that is part of a text frame (`para.framePr` set), per
 * ECMA-376 §17.3.1.11.
 *
 * The frame is OUT OF FLOW: it is drawn at an absolute (anchor-relative)
 * position and does NOT advance the in-flow `state.y`, so the following
 * non-frame paragraph begins where this paragraph sat. The frame's glyphs are
 * painted at their own run sizes (a drop cap's big letter is just a large `sz`
 * run, e.g. sample-11's 58.5 pt "D"). A wrap exclusion is then registered so
 * following body text flows around the frame.
 *
 * `anchorLineHpx` is one line height of the following non-frame paragraph,
 * needed to size a drop cap by `lines` (§17.3.1.11).
 */
function renderFrameParagraph(
  para: DocParagraph,
  state: RenderState,
  anchorLineHpx: number,
): void {
  const fp = para.framePr!;
  // In-flow Y the following paragraph must resume from. The frame is out of
  // flow; we restore this after drawing so state.y is untouched by the frame.
  const inFlowY = state.y;
  const box = resolveFrameBox(para, state, anchorLineHpx);

  // Draw the frame's glyphs by redirecting the flow geometry to the frame box,
  // then rendering the paragraph through the normal line path. inFrame=true
  // suppresses anchor-float re-registration and avoids re-entering the frame
  // dispatch; suppressSpaceBefore=true keeps the cap anchored to the paragraph
  // top (the frame is positioned absolutely, not in flow).
  const savedX = state.contentX;
  const savedW = state.contentW;
  state.contentX = box.x;
  state.contentW = Math.max(box.w, box.exRight - box.x);
  state.y = box.y;
  renderParagraph(para, state, true, undefined, /* inFrame */ true);
  state.contentX = savedX;
  state.contentW = savedW;

  // Restore the in-flow cursor: the frame consumes NO vertical space in the
  // body flow (§17.3.1.11 — the frame is positioned relative to the next
  // non-frame paragraph; the frame itself does not advance the flow).
  state.y = inFlowY;

  registerFrameFloat(box, fp, state);
}

interface NumberingMarkerLayout {
  numTab: number;
  picBullet: { bmp: DecodedImage; w: number; h: number } | null;
  numBodyOffset: number;
  markerJcShiftPx: number;
  /** Resolved marker ink width in px. Uses the decoded picture-bullet width when
   *  present, otherwise the marker glyph measurement in its resolved font. */
  markerWidthPx: number;
  markerTextLayout: NumberingMarkerTextLayout | null;
  hasMarker: boolean;
}

function resolveNumberingMarker(
  para: DocParagraph,
  state: RenderState,
  indLeft: number,
  indFirst: number,
): NumberingMarkerLayout {
  const { ctx, scale, fontFamilyClasses } = state;
  // Numbering marker. `hasMarker` is the "this paragraph has a marker" flag;
  // it is true for a text/glyph marker (`numMarker`) AND for a §17.9.9 picture
  // bullet (whose lvlText is typically empty — `numMarker` would be falsy).
  let numMarker = '';
  let numTab = 0;
  // §17.9.9/§17.9.20 — when the level uses a picture bullet, this holds its
  // decoded bitmap + draw size (px); the marker is drawn as an image, not text.
  let picBullet: { bmp: DecodedImage; w: number; h: number } | null = null;
  // First-line body offset (px) from paraX for an LTR numbered paragraph, set by
  // the §17.9.28 `<w:suff>` that follows the marker:
  //   tab (default) → body advances to the indentLeft tab stop (offset 0),
  //   space/nothing → body abuts the marker (marker end, + one space for space).
  let numBodyOffset = 0;
  // §17.9.8 `<w:lvlJc>` — horizontal shift (px) applied to the LTR marker draw so
  // it left/right/centre-aligns at the hanging-indent reference (firstLineX).
  // 0 = left (default); −markerW = right (period-aligned numerals: right edge at
  // firstLineX); −markerW/2 = centre. Set in the numbering block below.
  let markerJcShiftPx = 0;
  let markerWidthPx = 0;
  let markerTextLayout: NumberingMarkerTextLayout | null = null;
  if (para.numbering) {
    const markerShapeInput = numberingMarkerShapeInput(
      para.numbering,
      getDefaultFontSize(para),
    );
    numMarker = para.numbering.text;
    numTab = para.numbering.tab * scale;
    const suff = para.numbering.suff || 'tab';
    const pbPath = para.numbering.picBulletImagePath;
    if (pbPath) {
      const bmp = state.images.get(imageKey(pbPath));
      if (bmp) {
        picBullet = {
          bmp,
          w: (para.numbering.picBulletWidthPt ?? markerShapeInput.fontSizePt) * scale,
          h: (para.numbering.picBulletHeightPt ?? markerShapeInput.fontSizePt) * scale,
        };
      }
    }
    // Marker glyph width (px) with its RESOLVED font (§17.3.2.26 + §17.9.6); the
    // picture bullet's own width when present. Needed for both the suff≠tab abut
    // and the suff=tab overrun check below, so measure once up front.
    let markerW: number;
    if (picBullet) {
      markerW = picBullet.w;
    } else {
      markerTextLayout = shapeNumberingMarkerText(
        markerShapeInput,
        markerDisplayText(para.numbering),
        scale,
        state.layoutServices?.text,
      );
      if (markerTextLayout) {
        markerW = markerTextLayout.shape.advancePt;
      } else {
        // Only synthetic states that bypass the document service reach this
        // compatibility branch; production pagination and paint always share
        // the immutable service instance.
        ctx.font = buildFont(
          false, false, getDefaultFontSize(para) * scale,
          para.numbering.fontFamily ?? null, fontFamilyClasses,
        );
        markerW = ctx.measureText(markerDisplayText(para.numbering)).width;
      }
    }
    markerWidthPx = markerW;
    // §17.9.8 lvlJc: shift the marker so its left/right/centre aligns at
    // firstLineX (the hanging-indent reference). The marker's RIGHT edge measured
    // from paraX (the indentLeft tab) is then `indFirst + shift + markerW`.
    const lvlJc = para.numbering.jc || 'left';
    markerJcShiftPx = lvlJc === 'right' ? -markerW : lvlJc === 'center' ? -markerW / 2 : 0;
    const markerEndFromIndent = indFirst + markerJcShiftPx + markerW;
    if (suff !== 'tab') {
      const spaceW = suff === 'space'
        ? shapeNumberingMarkerText(
            markerShapeInput,
            ' ',
            scale,
            state.layoutServices?.text,
          )?.shape.advancePt ?? ctx.measureText(' ').width
        : 0;
      // body abuts the marker's right edge (+ one space for suff="space").
      numBodyOffset = markerEndFromIndent + spaceW;
    } else {
      // suff=tab: the marker is followed by a tab that advances the body to the
      // numbering's indentLeft tab stop (numBodyOffset 0 — the body sits at
      // paraX). But ECMA-376 §17.9.6 + §17.3.1.37: a tab never moves BACKWARD, so
      // when the marker overruns that stop — a wide multi-level number like
      // "1.1.1." whose glyphs exceed the hanging indent (the marker `indFirst`
      // budget), e.g. in a substitute font — the tab advances to the next stop
      // PAST the marker end instead, and the body follows it. Without this the
      // body stays at indentLeft and the marker overprints it (sample-11's
      // "1.1.1. Three" collided; Word advances "Three" to the next default tab).
      // markerEndFromIndent (jc-adjusted right edge from paraX) ≤ 0 ⇒ it fits
      // (right-aligned markers always do), leave the body at indentLeft.
      if (markerEndFromIndent > 0) {
        // Next tab stop strictly past the marker end, resolved in TEXT-MARGIN
        // coordinates via the SAME helper as line layout (§17.3.1.37 +
        // §17.15.1.25): honour the paragraph's explicit stops (already in margin
        // px = pos * scale) plus the document's automatic grid AFTER all custom
        // stops, then convert back to paraX-relative px (− indLeft).
        const markerEndFromMargin = indLeft + markerEndFromIndent;
        const customStopsPx = (para.tabStops ?? []).map((ts) => ({
          pos: ts.pos * scale,
          alignment: ts.alignment,
          leader: ts.leader,
        }));
        const stop = nextTabStop(markerEndFromMargin, customStopsPx, state.defaultTabPt * scale);
        if (stop) numBodyOffset = stop.pos - indLeft;
      }
    }
  }
  // True when the paragraph has any marker to draw (text glyph OR picture bullet).
  const hasMarker = numMarker !== '' || picBullet !== null;
  return {
    numTab,
    picBullet,
    numBodyOffset,
    markerJcShiftPx,
    markerWidthPx,
    markerTextLayout,
    hasMarker,
  };
}

/** Resolved numbering-marker bounds in the paragraph's physical coordinate
 * space. §17.9.7 applies lvlJc at the logical first-line margin; the RTL result
 * is therefore the physical mirror of LTR. `markerWidthPx` already represents
 * either measured text ink or the resolved picture-bullet width. */
function numberingMarkerBorderBounds(
  contentX: number,
  contentW: number,
  physicalIndentLeft: number,
  physicalIndentRight: number,
  firstIndent: number,
  markerJcShiftPx: number,
  markerWidthPx: number,
  numTab: number,
  baseRtl: boolean,
): { left: number; right: number } {
  if (!baseRtl) {
    const left = contentX + physicalIndentLeft + firstIndent + markerJcShiftPx;
    return { left, right: left + markerWidthPx };
  }
  const anchor = contentX + contentW - physicalIndentRight - firstIndent;
  const mirroredRight = anchor - markerJcShiftPx;
  const mirroredLeft = mirroredRight - markerWidthPx;
  // The established RTL draw path anchors the marker's right edge `numTab`
  // beyond the aligned body start. Its maximum physical-right position is the
  // paragraph start edge plus numTab (other text alignments can only move it
  // inward). Union that actual paint envelope with the mirrored lvlJc bounds.
  const paintedRight = contentX + contentW - physicalIndentRight + numTab;
  return {
    left: Math.min(mirroredLeft, paintedRight - markerWidthPx),
    right: Math.max(mirroredRight, paintedRight),
  };
}

function renderParagraph(
  para: DocParagraph,
  state: RenderState,
  suppressSpaceBefore = false,
  /** When set, render only `lines[start, end)` of the laid-out paragraph,
   *  used by the paginator to split paragraphs that don't fit on one page. */
  lineSlice?: { start: number; end: number; continues?: boolean },
  /** True when this call is the redirected draw of a `<w:framePr>` frame
   *  paragraph (from {@link renderFrameParagraph}). It suppresses the in-flow
   *  cursor bookkeeping that the frame path handles itself: anchor-float
   *  registration is skipped (the frame is the float) and the
   *  topAndBottom-skip / frame dispatch are bypassed (the geometry is already
   *  the frame box). Frame dispatch for a non-frame call lives in
   *  renderBodyElements so it can pass the anchor paragraph's line height. */
  inFrame = false,
  /** ECMA-376 §17.3.1.7 paragraph-border merge: suppress the top edge when a
   *  same-border paragraph precedes this one in the same column, and the bottom
   *  edge when one follows. Computed by the paint loop (renderBodyElements /
   *  renderParaList), which knows in-flow adjacency. Absent ⇒ draw the full box
   *  (a standalone bordered paragraph). */
  borderMerge?: ParaBorderMerge,
  /** PR 5 — pre-measured scale-1 line partition supplied by body fragment paint
   *  ({@link paintParagraphFragment}). When provided (even empty), the paragraph's
   *  lines are the SUPPLIED partition rescaled to the paint scale — the reuse gate,
   *  the scale-1 recompute, and the float re-layout are all bypassed. This makes
   *  paint consume stored geometry without re-running {@link layoutLines}. The
   *  paint pass is byte-identical to the legacy acquisition because the fragment
   *  holds exactly the scale-1 lines the legacy non-float path would compute
   *  (migration is gated to non-float, non-marker paragraphs). Empty ⇒ the
   *  markOnly / anchor-only paragraph, handled by the existing empty-mark branch. */
  suppliedScale1Lines?: readonly LayoutLine[],
): void {
  const { ctx, scale, contentX, contentW, defaultColor, dryRun, fontFamilyClasses } = state;
  const paragraphContext = resolveStateParagraphLayoutContext(state, para);
  // Capture Y before spaceBefore — used for paragraph-relative anchor image positioning
  const paragraphStartY = state.y;

  if (!suppressSpaceBefore) state.y += para.spaceBefore * scale;

  // Register anchor floats from this paragraph. ECMA-376 §20.4.3.5: a
  // `positionV relativeFrom="paragraph"` float is positioned relative to "the
  // paragraph which contains the drawing anchor" — its TOP edge, BEFORE the
  // paragraph's spaceBefore (Word anchors the float at the paragraph top, not the
  // post-spaceBefore text area). So pass `paragraphStartY` (pre-spaceBefore),
  // identically for wrap AND wrapNone floats (renderAnchorImages below already
  // uses paragraphStartY). Anchoring wrap floats at the post-spaceBefore text top
  // placed them spaceBefore too low — e.g. sample-12's figure (anchor paragraph
  // spaceBefore=12 pt) sat 12 pt under Word, eating the gap above its caption.
  // Skipped for the frame-draw recursion: a frame paragraph's wrap exclusion is
  // its own FloatRect (renderFrameParagraph), not an anchor image/shape float.
  if (!inFrame) registerAnchorFloats(para, state, paragraphStartY);

  // behindDoc shapes must render before text so they appear behind it.
  // Float registration above remains unconditional per-slice bookkeeping; only
  // the paragraph-level anchor DRAW is restricted to the original first slice.
  if (!lineSlice || (lineSlice.start === 0 && !lineSlice.continues)) {
    renderAnchorImages(para, state, paragraphStartY, 'behind');
  }

  // If any topAndBottom float already extends past state.y, skip past it before
  // text starts. Scoped to this paragraph's column band (§20.4.2.20 / §17.6.4):
  // a topAndBottom float anchored in another newspaper column must not push this
  // column's text down — state.floats is page-scoped across columns, and
  // state.contentX/contentW is this element's column band (set per column by the
  // paint loop).
  state.y = skipPastTopAndBottom(state.y, state.floats, contentX, contentX + contentW);

  const textAreaTopY = state.y;

  // ECMA-376 §17.3.1.12 w:ind — the transitional left/right attributes are
  // logical start/end (Part 4 §14.11.2). In a bidi paragraph the start side is
  // the physical RIGHT, so the two indents swap physical sides here.
  //
  // A frame paragraph's own body-style indents (e.g. a default first-line indent
  // inherited from the body style) do NOT apply to the frame content: the frame
  // box already positions the glyphs from its left edge, and the wrap exclusion
  // is built from that same left edge (§17.3.1.11). Honoring the indent here
  // would shift the cap glyph right of the exclusion band and let body text
  // overlap it, so zero the indents in the frame-draw recursion.
  const baseRtl = paragraphContext.baseRtl;
  const indLeft = inFrame ? 0 : paragraphContext.physicalIndentLeftPt * scale;
  const indRight = inFrame ? 0 : paragraphContext.physicalIndentRightPt * scale;
  const indFirst = inFrame ? 0 : para.indentFirst * scale;

  // Numbering marker layout (§17.9.x): see resolveNumberingMarker.
  const {
    numTab, picBullet, numBodyOffset, markerJcShiftPx, markerWidthPx, markerTextLayout, hasMarker,
  } =
    resolveNumberingMarker(para, state, indLeft, indFirst);

  const paraX = contentX + indLeft;
  const firstLineX = paraX + indFirst;
  const paraW = contentW - indLeft - indRight;
  const markerBounds = hasMarker
    ? numberingMarkerBorderBounds(
        contentX,
        contentW,
        indLeft,
        indRight,
        indFirst,
        markerJcShiftPx,
        markerWidthPx,
        numTab,
        baseRtl,
      )
    : undefined;
  const borderBox = paragraphBorderContentBox(
    contentX,
    contentW,
    indLeft,
    indRight,
    indFirst,
    baseRtl,
    markerBounds,
  );

  // ECMA-376 §17.9.28 (`<w:suff>`) governs where a numbering marker's first-line
  // body starts. With suff=tab (default) the body advances to the indentLeft tab
  // stop (`numBodyOffset`); §17.3.1.6 makes `<w:ind>` logical under `<w:bidi>`, so
  // this applies to the RTL body's start (physical-right) edge just as it does to
  // the LTR body's left edge.
  //
  // The RTL branch is gated to a genuine HANGING indent (`indFirst < 0`) — the only
  // shape a real numbered/bulleted list uses (§17.3.1.12: the marker sits in the
  // hanging margin). A non-hanging RTL marker (positive/zero first-line indent, a
  // degenerate authoring) keeps its legacy raw-`indFirst` handling so it stays
  // consistent with the measure pass (which cannot recompute `numBodyOffset`), and
  // suff=space/nothing (body abuts the marker) is likewise EXCLUDED — its RTL
  // mirror is a follow-up. LTR is unaffected by these RTL-only guards (it already
  // used numBodyOffset for every suffix and indent), so LTR stays byte-identical.
  const markerUsesBodyOffset =
    hasMarker
    && (!baseRtl || ((para.numbering?.suff || 'tab') === 'tab' && indFirst < 0));

  // Collect all text segments with formatting (resolving field runs against page context)
  const segments = buildSegments(para.runs, segmentEnvironmentOf(state));
  // Word renders ruby paragraphs with consistent line spacing — every line
  // in a paragraph that carries ANY furigana snaps to the same pitch
  // multiple. Compute once at paragraph scope and share with the line loop.
  const paraHasRuby = paragraphContext.hasRuby;
  const grid = paraGrid(para, state);

  // A paragraph with no inline content (literally empty, or anchor-only) still
  // produces ONE paragraph-mark line box (ECMA-376 §17.3.1.29 regulates only the
  // existence of that line; the horizontal wrap geometry around a square float is
  // §20.4.2.17). The displacement below — flow the mark line below the float band
  // when the side gap cannot hold the pilcrow — has no dedicated §x.x.x: the only
  // SPEC-mandated flow of a line onto a float-free region is the explicit
  // `<w:br w:clear>` of §17.18.3, which is not what fires here. The TRIGGER for an
  // EMPTY paragraph mark is Word's measured behaviour: the mark stays BESIDE the
  // float as long as the free side-gap can hold the pilcrow itself (its em width),
  // and drops below only when the gap is narrower than that — i.e. effectively a
  // full-width float band. This is NARROWER than the 1-inch rule Word applies to
  // CONTENT lines (issue #676, wordMinLineStartPx): an empty mark in a ~62pt gap
  // (under 1 inch) still sits beside the float — flowing it below at 1 inch pushed
  // sample-12's caption + CONCLUSION onto the next page (#676 over-generalized the
  // content-line threshold onto empty marks; this restores the pilcrow threshold).
  // Grounded from sample-9 p.4 (full-width band → drops below, carrying
  // its wrapNone anchor image) and sample-12 p.2 (~62pt gap → beside). Without the
  // drop-below an empty mark wedges into a sub-pilcrow sliver beside a full-width
  // float band and the following paragraphs (and any wrapNone image they anchor)
  // stay pinned inside the band. We resolve the mark line's flowed top here and
  // use it for the mark advance, the shading/border rect, and the
  // paragraph-relative base of any wrapNone anchor image drawn below.
  const markProbeHeight = paragraphMarkLineHeight(
    para,
    scale,
    grid,
    paraHasRuby,
    state.docEastAsian,
    ctx,
    state.fontFamilyClasses,
    para.lineSpacing,
    state.resolvedLocalFonts,
    state.layoutServices?.text,
    paragraphMarkShapeInput(para),
  );
  const resolveEmptyMarkTop = (): number => {
    if (state.floats.length === 0) return textAreaTopY;
    // Required side-gap for the mark line: the pilcrow's em width
    // (paragraphMarkEmPx) — the empty-mark threshold, NOT the 1-inch content-line
    // rule (issue #676). A gap narrower than the pilcrow cannot hold the
    // mark, so it flows below the band.
    const win = resolveLineFloatWindow(
      textAreaTopY, paragraphMarkEmPx(para, scale), markProbeHeight, paraX, paraW, state.floats,
      // Raw COLUMN band for the topAndBottom gate (§20.4.2.20 / §17.6.4): an
      // empty mark under a topAndBottom float in this column's indent margin
      // still flows below it, matching the measure pass (measureMarkOnly).
      contentX, contentX + contentW,
      {
        xLeftPt: 0,
        xRightPt: state.pageWidth * scale,
        readingDirection: baseRtl ? 'rtl' : 'ltr',
      },
    );
    return win.topY;
  };

  if (segments.length === 0) {
    // Literally-empty paragraph: one paragraph-mark line box, no inline content
    // and (by construction in the paginator) never sliced.
    renderEmptyMarkParagraph(para, state, {
      grid, paraHasRuby, contentX, indLeft, paraW,
      borderX: borderBox.x, borderW: borderBox.w,
      textAreaTopY, paragraphStartY,
      markTop: resolveEmptyMarkTop(), totalLines: 0, lineSlice: undefined, borderMerge,
    });
    return;
  }

  const wrapCtx: WrapLayoutCtx | undefined = state.floats.length > 0 ? {
    startPageY: state.y,
    paraX,
    // Raw COLUMN band for the topAndBottom gate (§20.4.2.20 / §17.6.4). `paraX`
    // above is the indented text band; the two diverge under a left indent, and
    // a topAndBottom float in this column's indent margin must still push text
    // below it. state.contentX/contentW is this element's column band (set per
    // column by the paint loop), matching the measure pass.
    columnXPt: contentX,
    columnWidthPt: contentW,
    referenceXPt: 0,
    referenceWidthPt: state.pageWidth * scale,
    readingDirection: baseRtl ? 'rtl' : 'ltr',
    floats: state.floats,
    paragraphMarkLineStartWidth: paragraphMarkEmPx(para, scale),
    lineBoxH: (a, d, _h, is, ea, gc) => lineBoxHeight(
      para.lineSpacing,
      a,
      d,
      scale,
      grid,
      paraHasRuby,
      is ?? 0,
      // §17.6.5 cell rounding follows this line's script, matching text boxes;
      // ruby paragraphs retain their established uniform paragraph resolver.
      paraHasRuby ? paragraphContext.hasEastAsianText : (ea ?? false),
      gc,
    ),
    pageH: state.pageH,
  } : undefined;
  // The canonical scale-1 paint bridge is intentionally limited to ordinary
  // horizontal paragraphs in the body story, including body-table cells (and
  // nested body tables). Float wrapping still lays out directly in paint-space;
  // non-body stories, frames, bidi/kashida, ruby, emphasis marks, and vertical
  // text retain their established scale-aware draw paths until each has dedicated
  // canonical-transform coverage.
  const canonicalTextScale = canonicalParagraphTextScaleEligible(
    state.storyContext ?? BODY_STORY_CONTEXT,
    state.verticalCJK,
    inFrame,
    wrapCtx !== undefined,
    paragraphContext,
    para,
    segments,
  );

  // ECMA-376 §17.3.1.12 (hanging) + §17.3.1.38 (a hanging indent implicitly
  // creates a tab stop at indentLeft) + §17.9.28 (`<w:suff>`, default "tab"):
  // in a hanging-indent list the number glyph sits at firstLineX (= indentLeft −
  // hanging); with suff=tab it is followed by a tab that advances the body to the
  // indentLeft tab stop, so the first line's TEXT region matches the continuation
  // lines' ([paraX, paraX+paraW]) and the negative first-line indent positions
  // only the marker. With suff=space/nothing the body abuts the marker instead
  // (numBodyOffset, computed above). Non-numbered paragraphs apply the first-line
  // indent (positive firstLine, or a bare negative hanging without a marker) to
  // the body as usual.
  //
  // §17.3.1.6 makes `<w:ind>` (and its hanging first-line component) logical under
  // `<w:bidi>`, so this whole construction is direction-symmetric: an RTL list
  // mirrors it to the physical RIGHT — the marker sits in the hanging margin past
  // the start (right) edge and the suff=tab body still starts at the indentLeft tab
  // stop. So a suff=tab marker uses `numBodyOffset` for BOTH directions
  // (`markerUsesBodyOffset`); the RTL start-edge placement then falls out of
  // `effAvailW` (the negative first-line indent must NOT widen the body, only
  // position the marker).
  const firstLineIndent = markerUsesBodyOffset ? numBodyOffset : firstLineX - paraX;
  const paintGridDeltaPx = gridCharDeltaPx(grid, scale);
  // Legacy non-retained stories still resolve their line partition in point
  // space and map it through the Canvas viewport. Body and retained table-cell
  // paragraphs bypass this branch and paint their acquired ParagraphLayout.
  const paraW1 = contentW / scale
    - (inFrame ? 0 : paragraphContext.physicalIndentLeftPt)
    - (inFrame ? 0 : paragraphContext.physicalIndentRightPt);
  const indLeft1 = inFrame ? 0 : paragraphContext.physicalIndentLeftPt;
  const firstIndent1 = markerUsesBodyOffset ? numBodyOffset / scale : para.indentFirst;
  const gridDelta1 = gridCharDeltaPx(grid, 1);
  // ECMA-376 §17.3.3.23 — paraX-relative X of the text-margin right edge, for
  // resolving a `<w:ptab w:relativeTo="margin">` (paraW is the content box; add
  // the right indent to reach the margin). Scale and scale-1 mirrors kept in sync.
  const indRight1 = inFrame ? 0 : paragraphContext.physicalIndentRightPt;
  const marginRightPx = paraW + indRight;
  const marginRightPx1 = paraW1 + indRight1;
  const lines = suppliedScale1Lines !== undefined
    // Body fragment paint: use the fragment's scale-1 partition, rescaled to the
    // paint scale via the canonical geometry bridge. No re-layout, so paint
    // scales stored geometry only (scale 1 returns the partition unchanged, so a
    // scale-1 paint invokes no measureText).
    ? rescaleLayoutLines([...suppliedScale1Lines], scale, ctx, state.fontFamilyClasses, paintGridDeltaPx, canonicalTextScale)
    : wrapCtx
      ? layoutLines(ctx, segments, paraW, firstLineIndent, scale, para.tabStops, wrapCtx, state.fontFamilyClasses, indLeft, state.kinsoku, paintGridDeltaPx, state.defaultTabPt, marginRightPx, baseRtl, jcIsFullyJustified(para.alignment), jcStretchesLastLine(para.alignment))
      : rescaleLayoutLines(
          layoutLines(ctx, segments, paraW1, firstIndent1, 1, para.tabStops, undefined, state.fontFamilyClasses, indLeft1, state.kinsoku, gridDelta1, state.defaultTabPt, marginRightPx1, baseRtl, jcIsFullyJustified(para.alignment), jcStretchesLastLine(para.alignment)),
          scale, ctx, state.fontFamilyClasses, paintGridDeltaPx, canonicalTextScale,
        );

  // Decimal-tab auto-alignment. ECMA-376 (§17.3.1.37 tabs / §17.18.84 ST_TabJc
  // `decimal`) only positions content at a tab stop when an explicit tab
  // character advances to it; absent a tab, content starts at the indent. Word,
  // however, aligns a bare number to a leading DECIMAL tab with NO tab character
  // — the built-in "Decimal Aligned" paragraph style on table number cells does
  // exactly this. sample-11's College table proves it: 110 / 103 / +7 etc. each
  // right-align on the decimal tab at 18 pt (Word PDF bbox: per-column right
  // edges coincide), where we previously left-aligned them. This is a deliberate
  // Word-runtime deviation (user-approved); it is gated to NUMERIC content whose
  // first tab stop is `decimal` and which carries no explicit tab, so ordinary
  // paragraphs are untouched. We right-edge align the number at the stop — the
  // same approximation the explicit-tab decimal path uses (frac=1; it does not
  // split on the '.'), exact for the integers in scope. Applied at DRAW time as
  // a pure horizontal offset, so the measured row height (and the paginate/paint
  // height contract) is unaffected.
  const decimalAutoTabPx: number | null = (() => {
    if (segments.some((s) => 'isTab' in s)) return null; // explicit tab wins
    const stops = para.tabStops ?? [];
    if (stops.length === 0) return null;
    const firstStop = stops.reduce((a, b) => (b.pos < a.pos ? b : a));
    if (firstStop.alignment !== 'decimal') return null;
    const txt = para.runs.map((r) => (r as { text?: string }).text ?? '').join('').trim();
    if (txt === '' || !/^[+\-(]?[\d., ]+\)?%?$/.test(txt)) return null; // numbers only
    return firstStop.pos * scale - indLeft; // px, relative to paraX (mirrors layoutLines' stopXof)
  })();

  // A paragraph whose only segments are wrap-float anchors (wp:anchor) places no
  // inline content on any line, so layoutLines returns zero lines. Per ECMA-376
  // §17.3.1.29 the paragraph mark still produces one line box; §20.4.2.x removes
  // the floating object from the inline flow but does not suppress that mark.
  // Reserve the same paragraph-mark line height the literal-empty path uses, so
  // consecutive anchor-only paragraphs don't collapse onto each other. The
  // anchor floats themselves are registered and drawn by registerAnchorFloats /
  // renderAnchorImages on their own absolute-position path, so this only adds the
  // in-flow paragraph-mark advance (no double counting, no double draw).
  if (lines.length === 0) {
    // Anchor-only paragraph: same content-less mark line as the literally-empty
    // path (the anchor floats themselves are drawn separately). Slice guards
    // honor a paginator-split slice (spaceAfter on the final slice, anchor
    // images on the first).
    renderEmptyMarkParagraph(para, state, {
      grid, paraHasRuby, contentX, indLeft, paraW,
      borderX: borderBox.x, borderW: borderBox.w,
      textAreaTopY, paragraphStartY,
      markTop: resolveEmptyMarkTop(), totalLines: lines.length, lineSlice, borderMerge,
    });
    return;
  }

  // For paragraphs that carry any ruby annotation, Word renders every line
  // at the SAME height. Per the user's note: when the section's docGrid is
  // active, Word widens the grid pitch to accommodate the tallest required
  // line (ruby + base + leading), then ALL lines in the paragraph use that
  // widened pitch — both ruby-bearing and non-ruby lines share the same
  // baseline grid, otherwise the lines drift. We mimic this by computing
  // uniformLineH = ceil(max natural / pitch) * pitch when docGrid is on,
  // else just the max natural.
  const uniformLineH = paraHasRuby
    ? snapParaLineToGrid(
        Math.max(0, ...lines.map(l => lineBoxHeight(para.lineSpacing, l.ascent, l.descent, scale, grid, true, l.intendedSingle, paragraphContext.hasEastAsianText))),
        grid,
        scale,
      )
    : 0;
  const lineHForLine = (l: typeof lines[number]): number =>
    paraHasRuby
      ? uniformLineH
      // §17.6.5 cell rounding is gated by the line's script; a Latin-only line
      // in a CJK paragraph keeps its natural height, matching the text-box path.
      : lineBoxHeight(para.lineSpacing, l.ascent, l.descent, scale, grid, false, l.intendedSingle, l.eastAsian ?? false, l.gridCountSingle);

  // Slice bounds — when the paginator split this paragraph across pages,
  // only render lines in [sliceStart, sliceEnd). The first line we paint
  // resets state.y baseline so the slice begins at the page's content top.
  // Resolved BEFORE the shading fill so the fill height covers exactly the
  // lines this page paints (see paintedParagraphHeight): a sliced paragraph
  // must not fill to the full-paragraph height past the slice's bottom border.
  const sliceStart = lineSlice ? lineSlice.start : 0;
  const sliceEnd = lineSlice ? lineSlice.end : lines.length;
  // The slice is authoritative for WHICH lines land on this page. Canonically
  // scaled body paragraphs retain the exact scale-1 count; excluded float/legacy
  // paths may still lay out at paint scale, so cap against their actual line array
  // and never index a phantom slice line. `paintEnd` also bounds the shading
  // height below.
  const paintEnd = Math.min(sliceEnd, lines.length);

  if (para.shading && !dryRun) {
    // Shading is the BACKGROUND (text paints on top), so its height must be known
    // BEFORE the draw loop. Replay the loop's exact per-line advancement over the
    // painted slice so the fill height === the post-loop border height
    // (state.y − textAreaTopY) BY CONSTRUCTION — the fill meets the bottom border
    // in the float-clearance and page-slice cases too, not just top/left/right.
    const paintedH = paintedParagraphHeight(lines, sliceStart, paintEnd, textAreaTopY, lineHForLine);
    ctx.fillStyle = `#${para.shading}`;
    const sb = paraShadingRect(borderBox.x, textAreaTopY, borderBox.w, paintedH, para.borders, borderMerge, scale);
    ctx.fillRect(sb.x, sb.y, sb.w, sb.h);
  }

  // ECMA-376 §17.18.44 ST_Jc: "both" / "justify" / "distribute" (and the kashida
  // + thaiDistribute variants) fully justify each line by expanding inter-word
  // spaces (and, for expansion, inter-CJK boundaries; thaiDistribute also opens
  // Thai grapheme-cluster boundaries). The last line is traditionally left-
  // aligned (not stretched) for "both"/kashida AND "thaiDistribute" (Word GT,
  // issue #959); only "distribute" also stretches the last line. The slack is
  // divided across the eligible gaps. (jc classification lives in bidi-line so the
  // §17.18.44 knowledge stays single-source.)
  const isJustified = jcIsFullyJustified(para.alignment);
  const stretchLastLine = jcStretchesLastLine(para.alignment);

  // Bidirectional text. The paragraph's base direction comes from w:bidi
  // (ECMA-376 §17.3.1.6). We engage the (exact) bidi pass only when the base is
  // RTL or the line actually contains strong-RTL characters, so pure-LTR
  // paragraphs keep their byte-identical fast path. `alignEdge` resolves
  // logical start/end against the base direction. (`baseRtl` is declared with
  // the indent swap above.)
  const paraNeedsBidi = baseRtl || segmentsHaveRtl(segments);
  const alignEdge = resolveAlignEdge(para.alignment, baseRtl);

  // ECMA-376 §17.6.5 character-grid delta (px per EA glyph) for the DRAW pass —
  // the SAME value layoutLines folded into measuredWidth. A pure-EA segment is
  // drawn so its glyphs occupy exactly `measuredWidth` (= natural + len·Δ): the
  // draw uses `justifiedPiecePositions(..., letterSpacingPx = Δ)`, whose final
  // glyph lands on the box edge, so the painted advance equals measuredWidth by
  // construction. See the gridCharDeltaPx / gridSegDeltaPx header.
  const drawGridDeltaPx = gridCharDeltaPx(grid, scale);
  const drawCtx: ParagraphLineDrawCtx = { ctx, scale, state, para, dryRun, defaultColor, fontFamilyClasses, contentX, contentW, lines, grid, paraHasRuby, paraX, firstLineX, paraW, indLeft, indFirst, continuesParagraph: lineSlice?.continues === true, baseRtl, hasMarker, markerUsesBodyOffset, numTab, numBodyOffset, markerJcShiftPx, markerWidthPx, markerTextLayout, picBullet, isJustified, stretchLastLine, alignEdge, paraNeedsBidi, decimalAutoTabPx, drawGridDeltaPx, canonicalTextScale, lineHForLine };
  for (let li = sliceStart; li < paintEnd; li++) {
    drawParagraphLine(li, drawCtx);
  }

  if (para.borders && !dryRun) {
    // `state.y` started this pass at `textAreaTopY` (captured above, untouched
    // until the draw loop) and the loop advanced it by exactly the per-line steps
    // `paintedParagraphHeight` replays (the topY float-clearance max-jump, then
    // `+= lineHForLine`). So `textH` here equals the `paintedH` the shading fill
    // used above — the fill meets this bottom border by construction, in the
    // normal, float-clearance and page-slice cases alike.
    const textH = state.y - textAreaTopY;
    drawParaBorders(ctx, borderBox.x, textAreaTopY, borderBox.w, textH, para.borders, scale, state.dpr, borderMerge);
  }

  // spaceAfter is paragraph-level; only emit it on the slice that covers
  // the FINAL line of the paragraph (or when no slice is set at all). §17.3.1.7: a
  // drawn bottom border extends `space + width/2` below the text box; reserve the
  // amount it pokes past spaceAfter (MAX) so the next paragraph clears the rule —
  // mirrors estimateParagraphHeight / renderEmptyMarkParagraph.
  const isFinalSlice = !lineSlice || lineSlice.end >= lines.length;
  if (isFinalSlice) {
    state.y += Math.max(para.spaceAfter, bottomBorderExtentPt(para.borders, borderMerge)) * scale;
  }

  // Anchor images are absolutely positioned — draw after inline flow.
  // Skip this for continuation slices: anchor positioning is paragraph-relative
  // and the first slice already painted them.
  if (!lineSlice || (lineSlice.start === 0 && !lineSlice.continues)) {
    renderAnchorImages(para, state, paragraphStartY);
  }
}

/** Per-line draw context for {@link drawParagraphLine}. Bundles the read-only
 *  paragraph-scope values the line loop reads (plus `state`, mutated by
 *  reference via `state.y`). Extracted from {@link renderParagraph} verbatim so
 *  the per-line drawing is a single thin call; no behaviour change. */
interface ParagraphLineDrawCtx {
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
  scale: number;
  state: RenderState;
  para: DocParagraph;
  dryRun: boolean;
  defaultColor: string;
  fontFamilyClasses: Record<string, string>;
  contentX: number;
  contentW: number;
  lines: LayoutLine[];
  grid: DocGridCtx;
  /** Any line in the paragraph carries a ruby annotation → every line uses the
   *  uniform (max-natural) box height and keeps the centred baseline grid, so the
   *  lineRule=auto "extra leading below" rule (§17.3.1.33, #990) is not applied. */
  paraHasRuby: boolean;
  paraX: number;
  firstLineX: number;
  paraW: number;
  indLeft: number;
  indFirst: number;
  continuesParagraph: boolean;
  baseRtl: boolean;
  hasMarker: boolean;
  markerUsesBodyOffset: boolean;
  numTab: number;
  numBodyOffset: number;
  markerJcShiftPx: number;
  markerWidthPx: number;
  markerTextLayout: NumberingMarkerTextLayout | null;
  picBullet: { bmp: DecodedImage; w: number; h: number } | null;
  isJustified: boolean;
  stretchLastLine: boolean;
  alignEdge: ReturnType<typeof resolveAlignEdge>;
  paraNeedsBidi: boolean;
  decimalAutoTabPx: number | null;
  drawGridDeltaPx: number;
  canonicalTextScale: boolean;
  lineHForLine: (l: LayoutLine) => number;
}

/** Draws line `li` of a paragraph. Extracted from {@link renderParagraph}'s
 *  per-line loop body verbatim (the loop simply calls this for each line);
 *  `state.y` is advanced by reference, exactly as before. */
function drawParagraphLine(li: number, c: ParagraphLineDrawCtx): void {
  const {
    ctx, scale, state, para, dryRun, defaultColor, fontFamilyClasses,
    contentX, contentW, lines, grid, paraHasRuby, paraX, firstLineX, paraW, indLeft,
    indFirst, continuesParagraph, baseRtl, hasMarker, markerUsesBodyOffset, numTab, numBodyOffset, markerJcShiftPx,
    markerWidthPx, markerTextLayout, picBullet, isJustified, stretchLastLine, alignEdge, paraNeedsBidi,
    decimalAutoTabPx, drawGridDeltaPx, canonicalTextScale, lineHForLine,
  } = c;
    const line = lines[li];
    // ECMA-376 §17.6.20 + Part 4 §14.11.7 (#988 re-adjudication): only an
    // UPRIGHT-vertical page (tbRl family) takes the per-glyph vertical draw
    // paths below. A `btLr` page (`state.verticalAllRotated`) is the horizontal
    // layout rotated wholesale — its glyphs draw through the ordinary
    // HORIZONTAL branches and ride the +90° page rotation, so CJK lies rotated
    // 90° CW exactly like Word's GT. Horizontal pages: false, byte-identical.
    const verticalUpright = !!state.verticalCJK && !state.verticalAllRotated;
    // First-line indent and numbering prefix only apply to the paragraph's
    // ORIGINAL first line, not the first line of a continuation slice.
    const firstLine = li === 0 && !continuesParagraph;
    // Last-line justification flips off only at the paragraph's true end —
    // mid-paragraph slices keep justifying through to the slice boundary.
    const isLastLine = li === lines.length - 1;

    // Honor wrap-computed line topY (may push past topAndBottom floats).
    if (line.topY !== undefined && line.topY > state.y) state.y = line.topY;

    // Baseline placement inside the line box. §17.3.1.33 defines the box SIZE
    // (line/lineRule); the placement WITHIN the box below is Word's OBSERVED
    // behaviour, measured against its PDF export (#990), not an ECMA rule:
    //   • lineRule="auto" is MULTIPLE spacing — `w:line` is a 240ths-of-a-line
    //     multiplier. Word pins the baseline at the natural ascent from the box
    //     top and places the multiplier's extra leading ENTIRELY BELOW the
    //     glyphs; it does NOT centre the natural line in the enlarged box.
    //     (Measured in the #981/#990 follow-up: a 2.0× 48pt title was displaced
    //     ~22pt when centred.) The substituted-font single-line FLOOR
    //     (intendedSingle) still centres the glyph box within the single design
    //     line (half-leading), so a Meiryo single-spaced line is byte-for-byte
    //     unchanged. A sub-single multiplier (0 < value < 1) compresses the
    //     entire line box, so the glyph box is centred across that shorter
    //     advance. This lets the ink protrude equally above and below instead of
    //     shifting it wholly toward the following row/content.
    //   • exact/atLeast keep the extra space split half above / half below
    //     (centred). docGrid line rules snap to whole cells and ruby paragraphs
    //     use a uniform box — both have their own within-box placement, so they
    //     keep the full-box centring too.
    // Draw-only: lineHForLine (the box advance) is identical either way, so the
    // line pitch and pagination are unaffected — this pins only WHERE the glyphs
    // sit inside the box. (The trailing-empty-mark page-fit reads a SEPARATE
    // centred below-baseline extent that is deliberately left unchanged; see
    // lineBelowBaselinePx in line-layout.ts.)
    const lineH = lineHForLine(line);
    // A floating-anchor host can reserve line/grid height while painting no
    // glyph. Keep that reservation in `lineH`, but position visible ink from
    // the segments that actually draw. This preserves the host paragraph's
    // advance without letting a differently-sized zero-ink run lower adjacent
    // caption text (Word-compatible mixed-anchor behavior).
    const visibleAscent = line.visibleAscent ?? line.ascent;
    const visibleDescent = line.visibleDescent ?? line.descent;
    const visibleIntendedSingle = line.visibleIntendedSingle ?? line.intendedSingle;
    const glyphNatural = visibleAscent + visibleDescent;
    const autoMultiple =
      para.lineSpacing?.rule === 'auto' && !paraHasRuby && !isGridLineRule(grid);
    // For auto multiple spacing at or above 1×, centre the glyph only within
    // the single design-line box (= max(glyphNatural, intendedSingle), matching
    // lineBoxHeight's `natural`); the multiplier's extra then falls below. Below
    // 1×, centre against the authored compressed line box itself.
    const compressedAuto = autoMultiple && (para.lineSpacing?.value ?? 1) < 1;
    const centerBox = autoMultiple && !compressedAuto
      ? Math.max(glyphNatural, visibleIntendedSingle)
      : lineH;
    const baseline = state.y + (centerBox - glyphNatural) / 2 + visibleAscent;

    // Per-line X range (may be narrower than paraW when wrapping around floats).
    const lineLeft = paraX + line.xOffset;
    const lineAvailW = line.availWidth;
    // First-line indent shifts the START edge: physical left for LTR; for RTL
    // the start is the right edge, so it narrows/widens the line's available
    // width instead of moving x (effAvailW below).
    // For a numbered first line (LTR) the body sits at lineLeft + numBodyOffset
    // (the indentLeft tab stop for suff=tab → offset 0; the marker end for
    // space/nothing); indFirst only pulls the marker into the hanging margin
    // (drawn below). Non-numbered first lines apply indFirst to the body directly.
    let x = firstLine && !baseRtl ? (hasMarker ? lineLeft + numBodyOffset : lineLeft + indFirst) : lineLeft;
    // RTL first-line width. For a bare (non-marker) indent the raw first-line
    // indent narrows (positive firstLine) or widens (hanging) the line, so the
    // body's start (right) edge tracks the indent — mirror of the LTR x-shift.
    // But a suff=tab numbering marker follows §17.9.28: the negative hanging indent
    // positions only the marker, and the body starts at the indentLeft tab stop, so
    // the first line's text region equals the continuation lines' — use
    // `numBodyOffset` (0 for suff=tab), NOT the raw `indFirst`, so the body does NOT
    // hang one `hanging` past the start edge. `markerUsesBodyOffset` also excludes
    // the suff=space/nothing RTL case (kept on legacy `indFirst` here — out of
    // scope), keeping this consistent with the `firstLineIndent` used for breaking.
    const effAvailW = baseRtl && firstLine
      ? lineAvailW - (markerUsesBodyOffset ? numBodyOffset : indFirst)
      : lineAvailW;

    // Visual draw order. Under bidi we reorder the line's segments per UAX#9
    // (rule L2) and draw each with ctx.direction matching its resolved
    // direction; ctx.textAlign stays physical 'left' so x is always the
    // segment's left edge. The LTR fast path is untouched (visual === null).
    // Computed before justification so the stretch bookkeeping below can use
    // the same (visual) domain as the draw loop.
    const visual: LineVisualOrder | null = paraNeedsBidi
      ? computeLineVisualOrder(line.segments, baseRtl)
      : null;
    if (paraNeedsBidi) ctx.textAlign = 'left';
    const segCount = line.segments.length;
    // The visually-last segment (its trailing edge is the line's physical end, so
    // no gap opens there). Consumed ONLY on the bidi path of the justification
    // distribution below: an LTR line excludes no segment (the kernel's content-
    // span trim already closes the final glyph's gap). Equals the logical last in
    // the LTR fast path.
    const lastDrawnSi = visual ? visual.order[segCount - 1] : segCount - 1;

    const lineWidth = line.segments.reduce((s, seg) => s + seg.measuredWidth, 0);
    const lineSlack = effAvailW - (x - lineLeft) - lineWidth;
    // §17.18.44: a `both` line justifies UNLESS it is the paragraph's true last
    // line OR ends at a manual `<w:br/>` (§17.3.3.1) — both terminate a logical
    // line and are left-aligned. `distribute` (stretchLastLine) still spreads
    // every line, including these.
    const endsLogicalLine = isLastLine || (line.endsWithBreak ?? false);
    const applyJustify = isJustified && (!endsLogicalLine || stretchLastLine);

    // Slack distribution across the line's gaps (§17.18.44). `segStretch` /
    // `distPerGap` drive the draw loop below; they are set either by the JUSTIFY
    // block (expansion / compression of a jc=both/distribute line) further down,
    // or — for a NON-justified line whose natural width overran the box because
    // layoutLines' fit judgment spent the shrink budget to keep it on one row —
    // by the compression here. The two are mutually exclusive (applyJustify vs
    // not), so there is no double distribution.
    let segStretch: Map<number, SegStretch> | null = null;
    let distPerGap = 0;
    let kashidaPlan: Map<number, KashidaSegmentPlan> | null = null;
    // First content segment in reading order. Leading-whitespace segments before
    // it (a paragraph's 字下げ indent) are NOT stretched — Word keeps the indent
    // fixed and distributes slack only across the line content (§17.18.44). Only
    // meaningful for LTR: under bidi the logical-leading segment is not the
    // visually-leading one, so leave the skip off (0) there.
    let firstContentSi = 0;
    if (!paraNeedsBidi) {
      for (let i = 0; i < segCount; i++) {
        const seg = line.segments[i];
        if (!('text' in seg) || /\S/.test((seg as LayoutTextSeg).text)) { firstContentSi = i; break; }
      }
    }
    // Shrink-to-fit compression for a NON-justified line that overflows the box
    // (lineSlack < 0). layoutLines placed the whole line here on the promise that
    // its inter-word spaces would be squeezed by up to SPACE_SHRINK_RATIO (its fit
    // test admits Δ ≤ ratio·Σspace); reproduce that squeeze so the last glyph lands
    // inside the box instead of overrunning its clip (sample-10 p1's centred title
    // "…Conference" — the final "e" was clipped). Same spaces-only mechanism the
    // justified negative-slack path uses. `shrinkDelta` (≤ 0) is folded into the
    // alignment slack below so the now-narrower line re-centres/re-aligns correctly.
    let shrinkDelta = 0;
    if (!applyJustify && lineSlack < 0) {
      const distSegs = line.segments.map(seg =>
        // §17.3.2.14 fixes fit-region pitch; §17.18.44 must therefore treat the
        // region like a non-text object so none of its internal gaps get slack.
        'text' in seg && (seg as LayoutTextSeg).fitTextRegionIndex === undefined
          ? { text: (seg as LayoutTextSeg).text }
          : {},
      );
      const shrinkDist = shrinkFitCompression(
        distSegs,
        lineSlack,
        firstContentSi,
        paraNeedsBidi ? lastDrawnSi : segCount,
        line.ascent,
      );
      if (shrinkDist) {
        segStretch = shrinkDist.perSeg;
        distPerGap = shrinkDist.perGap;
        shrinkDelta = distributedDelta(shrinkDist);
      }
    }
    // Alignment slack AFTER any shrink squeeze: the drawn width is
    // lineWidth + shrinkDelta, so the remaining slack the align offset centres /
    // right-aligns against is lineSlack − shrinkDelta (0 when the squeeze fully
    // absorbed the overflow ⇒ the line fills the box and align offset is 0).
    const alignSlack = lineSlack - shrinkDelta;
    let alignOffset = 0;
    // ECMA-376 §22.1.2.88 `m:jc` / §22.1.2.30 `m:defJc` — a display equation's
    // justification is independent of the paragraph's text alignment. When this
    // line is exactly one display-math segment, resolve its effective math jc
    // (per-instance → document default → spec default `centerGroup`) and use it
    // for THIS line only, overriding the paragraph alignEdge.
    const onlyMathSeg =
      line.segments.length === 1 &&
      'mathNodes' in line.segments[0] &&
      (line.segments[0] as LayoutMathSeg).display
        ? (line.segments[0] as LayoutMathSeg)
        : null;
    const mathEdge = onlyMathSeg
      ? mathJcToEdge(onlyMathSeg.jc ?? state.mathDefJc ?? 'centerGroup')
      : null;
    const effEdge = mathEdge ?? alignEdge;
    if (effEdge === 'right') {
      alignOffset = alignSlack;
    } else if (effEdge === 'center') {
      alignOffset = alignSlack / 2;
    } else if (effEdge === 'justify' && baseRtl && !applyJustify) {
      // The unstretched (last) line of a justified RTL paragraph aligns to the
      // leading edge — the RIGHT margin (§17.18.44 `both`: last line is
      // start-aligned). LTR keeps alignOffset 0 as before.
      alignOffset = alignSlack;
    }
    // 'left' and stretched 'justify' keep alignOffset 0.
    // Decimal-tab auto-alignment (see decimalAutoTabPx above): override the
    // paragraph alignment so the number's right edge (its decimal point, for an
    // integer) lands on the decimal tab. `paraX + decimalAutoTabPx` is the stop
    // in device space; subtracting the line width and the current `x` yields the
    // left-shift, clamped ≥ 0 so a number wider than the stop simply overflows
    // right (never pulled left of its natural start).
    if (decimalAutoTabPx != null && lineWidth > 0) {
      alignOffset = Math.max(0, paraX + decimalAutoTabPx - lineWidth - x);
    }
    x += alignOffset;

    if (firstLine && hasMarker && !dryRun) {
      if (picBullet) {
        // §17.9.9/§17.9.20 — the marker is an image. It occupies the same
        // anchor a text marker would (LTR: left edge in the hanging margin;
        // RTL: right edge numTab past the start edge), and rides the line's jc
        // alignment via `x`/`lineWidth` exactly like the glyph marker below.
        // Vertically it bottom-aligns to the baseline (the inline-image
        // convention, §17.3.3 anchored to the text bottom) so a sub-em bullet
        // rests on the line like a glyph.
        const { bmp, w, h } = picBullet;
        const top = baseline - h;
        // LTR: left edge at firstLineX, shifted by lvlJc (§17.9.8); RTL keeps its
        // own mirrored anchor.
        const left = baseRtl ? x + lineWidth + numTab - w : lineLeft + indFirst + markerJcShiftPx;
        ctx.drawImage(bmp, left, top, w, h);
      } else {
        const fallbackFontSize = getDefaultFontSize(para) * scale;
        // Marker ink (§17.9.24 + §17.3.1.29): the level rPr's own color wins;
        // absent that, Word layers the level rPr over the PARAGRAPH MARK's run
        // properties, so the mark's resolved color tints the bullet/number;
        // else the default ink. An EXPLICIT `w:color w:val="auto"` on the
        // level (colorAuto, §17.3.2.6) breaks that mark fallback — auto is a
        // named automatic color, not "unset" — and lands on the default ink.
        // Body-run colors never reach the marker (§17.9.24: the level rPr
        // "affects only the numbering text itself, not the remainder of runs
        // in the numbered paragraph").
        const markerColor = para.numbering!.color
          ?? (para.numbering!.colorAuto ? null : para.paragraphMarkColor);
        ctx.fillStyle = markerColor ? `#${markerColor}` : defaultColor;
        if (baseRtl) {
          // The RTL list marker is laid out INLINE at the line's start (right)
          // edge: its right edge sits numTab (w:hanging) to the right of the
          // text's start edge, mirroring the LTR `firstLineX - numTab` anchor,
          // and it follows the text through jc alignment (sample-8 PDF ground
          // truth: marker right edge = aligned text right edge + hanging).
          const prevAlign = ctx.textAlign;
          const prevDir = ctx.direction;
          ctx.textAlign = 'left';
          ctx.direction = 'rtl';
          const markerW = markerTextLayout?.shape.advancePt ?? markerWidthPx;
          const markerX = x + lineWidth + numTab - markerW;
          if (markerTextLayout) {
            paintNumberingMarkerText(ctx, markerTextLayout, markerX, baseline);
          } else {
            ctx.font = buildFont(
              false, false, fallbackFontSize,
              para.numbering!.fontFamily ?? null, fontFamilyClasses,
            );
            ctx.fillText(markerDisplayText(para.numbering!), markerX, baseline);
          }
          ctx.textAlign = prevAlign;
          ctx.direction = prevDir;
        } else {
          // Marker sits in the hanging margin at lineLeft + indFirst (= firstLineX
          // when the line isn't shifted by a float; lineLeft already includes any
          // float xOffset, so the marker tracks the body that hangs off it),
          // shifted by lvlJc (§17.9.8) so a "right" level period-aligns its right
          // edge at firstLineX. The body was advanced past the marker above
          // (numBodyOffset).
          const markerX = lineLeft + indFirst + markerJcShiftPx;
          if (markerTextLayout) {
            // §17.6.20 (tbRl) uses the same retained scalar routes; each routed
            // span advances from the geometry that positioned the marker.
            paintNumberingMarkerText(
              ctx,
              markerTextLayout,
              markerX,
              baseline,
              verticalUpright
                ? (paintCtx, text, drawX, drawBaseline, fontSizePx) => {
                    drawVerticalRun(paintCtx, text, drawX, drawBaseline, fontSizePx, 0);
                  }
                : undefined,
            );
          } else {
            ctx.font = buildFont(
              false, false, fallbackFontSize,
              para.numbering!.fontFamily ?? null, fontFamilyClasses,
            );
            ctx.fillText(markerDisplayText(para.numbering!), markerX, baseline);
          }
        }
        // Restore the default ink: everything after the marker previously ran
        // with fillStyle === defaultColor, and fills that don't set their own
        // style must keep seeing it.
        ctx.fillStyle = defaultColor;
      }
    }

    // Justified-line slack distribution (ECMA-376 §17.18.44). Positive slack
    // (lineWidth < availW) expands the line to fill the margin; negative slack
    // (lineWidth > availW, typically from canvas measuring ~1px wider than Word)
    // compresses it so the final glyph lands on the right margin instead of
    // overflowing. Gaps open at inter-word spaces AND — for expansion — inter-CJK
    // boundaries, so a pure-CJK `both`/`distribute` line fills the margin too
    // (Word fills CJK `both` lines by adding inter-character pitch; see
    // text-distribute.ts). distributeLineSlack returns, per logical segment, the
    // internal split points and a trailing-gap flag; the draw loop applies
    // `perGap` at each. Only computed when the line is a justify candidate
    // (jc=both/distribute, not the last line unless distribute) — a NON-justified
    // overflowing line was already squeezed above (`shrinkFitCompression`), and
    // `segStretch` / `distPerGap` / `firstContentSi` are hoisted before the align
    // offset so both distributions feed the SAME draw loop.
    if (applyJustify) {
      const slack = effAvailW - (x - lineLeft) - lineWidth;
      // Compression cap (negative slack): never eat more than ~a quarter em per
      // gap, estimated from the line ascent. For expansion this is unbounded.
      const minPerGap = -line.ascent * 0.25;
      // Expansion opens inter-CJK boundaries; compression touches only spaces
      // (shrinking a space is fine, overlapping ideographs is not).
      const distSegs = line.segments.map(seg =>
        // §17.3.2.14 fixes fit-region pitch; §17.18.44 must therefore treat the
        // region like a non-text object so none of its internal gaps get slack.
        'text' in seg && (seg as LayoutTextSeg).fitTextRegionIndex === undefined
          ? { text: (seg as LayoutTextSeg).text }
          : {},
      );
      // ECMA-376 §17.18.44 low/medium/highKashida first elongate valid
      // Arabic joins. Only the residual goes through the ordinary space/CJK
      // distributor; a line with no eligible join falls back to the full-slack
      // `both` behaviour. Upright-vertical (tbRl) text keeps the established
      // stage-1 path; an all-rotated (btLr) page follows the horizontal rules
      // like every other horizontal-branch behavior (measure == paint).
      const kashidaLevel = !verticalUpright
        ? kashidaLevelOf(para.alignment)
        : null;
      const kashidaDist = kashidaLevel
        ? computeLineKashidaDistribution(
            ctx,
            line.segments,
            slack,
            kashidaLevel,
            scale,
            fontFamilyClasses,
            drawGridDeltaPx,
          )
        : null;
      if (kashidaDist) kashidaPlan = kashidaDist.perSeg;
      const residualSlack = kashidaDist?.residualPx ?? slack;
      const dist = distributeLineSlack(
        distSegs,
        residualSlack,
        firstContentSi,
        // §17.18.44 spreads the slack across EVERY inter-CJK boundary on the line,
        // so the visually-last segment must still distribute pitch INTERNALLY when
        // it is a multi-glyph CJK run — else it stays at the bare grid pitch while
        // earlier segments absorb all the slack (two pitches on one line). The
        // kernel already closes only the FINAL glyph's gap via its content-span
        // trim, so an LTR line excludes no segment: pass `segCount`, the no-match
        // sentinel the pptx justifier also uses. Bidi keeps the whole-segment
        // exclusion because its logical-last unit is not the visually-last glyph.
        // See the `lastDrawnSi` option doc in core/src/text/line-distribute.ts.
        paraNeedsBidi ? lastDrawnSi : segCount,
        minPerGap,
        residualSlack > 0,
        // §17.18.44 thaiDistribute: on expansion, also open a gap at every Thai/
        // Lao/Khmer grapheme-cluster boundary so a space-free SEA line justifies
        // by inter-cluster pitch (Word GT: issue #959). `both`/`distribute` don't.
        para.alignment === 'thaiDistribute' && residualSlack > 0,
      );
      segStretch = dist ? dist.perSeg : null;
      distPerGap = dist ? dist.perGap : 0;
    }

    // ECMA-376 §17.3.2.4 (`<w:bdr>`): adjacent runs whose border attribute set
    // is identical form ONE run-border group and are "rendered within the same
    // set of borders". Accumulate the group's pixel extent as segments are
    // drawn left-to-right and stroke a single frame when the group ends (a
    // segment with a different / absent border, or end of line). Grouping is by
    // visual adjacency within this line; mixed-direction lines are an edge case
    // (the spec phrases the group in logical order) left for a follow-up.
    interface OpenBorderGroup {
      border: DocxRunBorder;
      left: number; right: number; top: number; bottom: number;
    }
    let borderGroup: OpenBorderGroup | null = null;
    const flushBorderGroup = () => {
      if (!borderGroup) return;
      const g = borderGroup;
      borderGroup = null;
      const bw = Math.max(1, g.border.width * scale); // w:sz/8 is in pt
      const sp = (g.border.space ?? 0) * scale;
      ctx.strokeStyle = g.border.color ? `#${g.border.color}` : defaultColor;
      ctx.lineWidth = bw;
      ctx.strokeRect(
        g.left - sp,
        g.top - sp,
        g.right - g.left + 2 * sp,
        g.bottom - g.top + 2 * sp,
      );
    };

    for (let vi = 0; vi < segCount; vi++) {
      const si = visual ? visual.order[vi] : vi;
      const seg = line.segments[si];
      if (visual) ctx.direction = visual.rtl[si] ? 'rtl' : 'ltr';
      // A non-text segment (tab / inline image / math) breaks run-border
      // adjacency (§17.3.2.4 groups adjacent *runs*), so close any open frame.
      if (!('text' in seg)) flushBorderGroup();
      if ('isTab' in seg) {
        // Tabs render as blank space, optionally filled with a leader (TOC dots etc.).
        if (!dryRun && seg.leader && seg.leader !== 'none' && seg.measuredWidth > 1) {
          drawTabLeader(ctx, seg.leader, x, baseline, seg.measuredWidth, seg.fontSize * scale, defaultColor, seg.bold, seg.italic);
        }
        x += seg.measuredWidth;
        continue;
      }
      if ('imagePath' in seg) {
        if (!dryRun) renderInlineImage(ctx, seg as LayoutImageSeg, x, baseline, scale, state.images, !!state.verticalCJK);
        x += seg.measuredWidth;
        continue;
      }
      if ('mathNodes' in seg) {
        const resourceKey = seg.mathResourceKey;
        const metadata = state.layoutServices?.math.resolve(resourceKey);
        const drawable = metadata?.available === false || !state.layoutServices
          ? undefined
          : privateResourceLookupOf<CanvasImageSource>(state.layoutServices)?.resolve(resourceKey);
        if (!dryRun && metadata && metadata.available !== false && drawable) {
          const emPx = seg.fontSize * scale;
          const w = metadata.widthEm * emPx;
          const h = (metadata.ascentEm + metadata.descentEm) * emPx;
          const top = baseline - metadata.ascentEm * emPx;
          ctx.drawImage(drawable, x, top, w, h);
        } else if (!dryRun && seg.fallbackText) {
          ctx.font = buildFont(false, false, seg.fontSize * scale, null, fontFamilyClasses);
          ctx.fillStyle = seg.color ?? defaultColor;
          ctx.fillText(seg.fallbackText, x, baseline);
        }
        x += seg.measuredWidth;
        continue;
      }
      const s = seg as LayoutTextSeg;
      const kashida = kashidaPlan?.get(si);
      const drawText = kashida?.text ?? s.text;
      // Justification stretch for THIS segment (logical index si). `internalStretch`
      // is the px added between the segment's own glyphs (inter-CJK boundaries);
      // `splitBefore` lists the code-point offsets to advance `distPerGap` at while
      // drawing. `spanW` (the glyph advance + internalStretch) covers every glyph
      // and the interior pitch; `decoW` (below) additionally covers the segment's
      // own widened trailing SPACE so run decorations stay gap-free under `both`.
      // §17.3.2.14 fitText is already a fixed-width cell; paragraph
      // justification must not stretch its internal glyph gaps a second time.
      const distributedStretch = segStretch?.get(si);
      // An augmented Arabic word must stay in one fillText so the browser keeps
      // contextual shaping. Ignore any residual distributor splitBefore points
      // on that segment; trailing-gap ownership remains valid and is read below.
      const stretch =
        !kashida && s.fitTextRegionIndex === undefined
          ? distributedStretch
          : undefined;
      // A fit region contributes an opaque atom to §17.18.44 distribution. Its
      // INTERNAL pitch stays suppressed above, but a legal boundary AFTER that
      // atom is paragraph slack, not §17.3.2.14 fit pitch, and must still advance
      // the following segment.
      const trailingDistributionGap = distributedStretch?.trailingGap ?? false;
      const internalStretch = (stretch?.internalStretch ?? 0) + (kashida?.advanceDeltaPx ?? 0);
      if (!dryRun) {
        const useCanonicalTransform = canonicalTextScale && scale !== 1;
        const effSizePx = calcEffectiveFontPx(s, scale);
        const glyphSizePx = calcEffectiveFontPx(s, useCanonicalTransform ? 1 : scale);
        // ECMA-376 §17.3.2.24 `<w:position>` — baseline raise(+)/lower(−) in pt.
        // Canvas y grows DOWNWARD, so a positive (raised) position subtracts from
        // y. It layers ON TOP of the super/sub offset (a positioned superscript
        // moves by both) and, per spec, does NOT change the font size or line box.
        const positionOffset = -(s.position ?? 0) * scale;
        const yOffset =
          (s.vertAlign === 'super'
            ? -s.fontSize * scale * 0.35
            : s.vertAlign === 'sub'
              ? s.fontSize * scale * 0.15
              : 0) + positionOffset;
        ctx.font = buildFont(s.bold, s.italic, glyphSizePx, s.fontFamily, fontFamilyClasses, s.fontRoute);

        // ECMA-376 §17.3.2.43 `<w:w>` horizontal glyph scale (1 = 100%) and
        // §17.3.2.35 `<w:spacing>` per-code-point character pitch in px. Both were
        // already folded into `s.measuredWidth` during layout, so decorations
        // (which use `decoW`/`spanW` below) follow automatically; here they drive
        // the glyph draw so paint == measure. §17.3.2.19 `<w:kern>` sets
        // `ctx.fontKerning` to match the measure pass exactly (see line-layout's
        // `setSegKerning`); restored after the glyph block.
        const segCharScale = s.charScale ?? 1;
        const segCharSpacingPx = s.fitTextPerGapPx ?? (s.charSpacing ?? 0) * scale;
        const prevFontKerning = ctx.fontKerning;
        if (s.kerning != null) {
          ctx.fontKerning = s.fontSize >= s.kerning ? 'normal' : 'none';
        }

        // Width spanned by the glyphs after justification, for ruby centring /
        // onTextRun reporting.
        const spanW = s.measuredWidth + internalStretch;

        // Width spanned by EVERY run decoration (highlight §17.3.1.15, shading
        // §17.3.2.32, border §17.3.2.4, underline §17.3.2.40, strike §17.3.2.37).
        // On a `both`/`distribute` line (§17.18.44) the justifier widens this
        // segment's TRAILING SPACE by `distPerGap` and advances the pen past it
        // (`x += distPerGap`, below). That space belongs to THIS run, so Word
        // paints its decorations across the widened advance — otherwise a GAP
        // opens between words (the bug this fixes). We extend ONLY when the
        // segment actually ends in whitespace it owns: an inter-CJK boundary gap
        // (no trailing space, e.g. a run/script split between two ideographs) is
        // NOT owned by either run, so extending there would bleed a highlight
        // past its run. Disabled under bidi: the pen advances in visual order
        // while `trailingGap` is a logical-order flag, so the widened gap is not
        // reliably the segment's own physical-right edge (kept at `spanW`, the
        // pre-justify-slack behaviour — no regression, just no slack fill).
        const ownsTrailingSlack =
          !!stretch?.trailingGap && !paraNeedsBidi && /\s$/.test(s.text);
        const decoW = spanW + (ownsTrailingSlack ? distPerGap : 0);

        // Glyph box used by every run-level box decoration (highlight fill,
        // §17.3.2.32 shading fill, §17.3.2.4 border): same vertical extent of
        // ~0.85em above the baseline to ~0.25em below it. Computed once so the
        // three decorations stay byte-identical (no duplicated 0.85 / 1.1).
        const boxTop = baseline + yOffset - effSizePx * 0.85;
        const boxHeight = effSizePx * 1.1;

        if (s.highlight) {
          ctx.fillStyle = HIGHLIGHT_COLORS[s.highlight] ?? '#FFFF00';
          ctx.fillRect(x, boxTop, decoW, boxHeight);
        }

        // ECMA-376 §17.3.2.32 run shading fill (`<w:shd w:fill>`): a solid
        // background rect behind the glyphs. Used for inverse video (black fill
        // + automatic = white text). Same rect geometry as the highlight box.
        if (s.background) {
          ctx.fillStyle = `#${s.background}`;
          ctx.fillRect(x, boxTop, decoW, boxHeight);
        }

        // ECMA-376 §17.3.2.4 run border (`<w:bdr>`, "box"): a rectangle around
        // the run, inflated by w:space (pt → px), drawn after the background so
        // the box outlines the filled area. Per the spec, adjacent runs sharing
        // an identical border render within the SAME frame, so instead of
        // stroking here we extend (or open) the current border group; the frame
        // is stroked by flushBorderGroup() when the group ends. The box bounds
        // each segment's glyph box (same rect the shading uses) unioned across
        // the group, so a mixed-size group still encloses every run.
        const activeBorder =
          s.border && s.border.style !== 'none' && s.border.style !== 'nil'
            ? s.border
            : null;
        if (activeBorder) {
          const segTop = boxTop;
          const segBottom = segTop + boxHeight;
          if (borderGroup && runBordersEqual(borderGroup.border, activeBorder)) {
            borderGroup.right = x + decoW;
            borderGroup.top = Math.min(borderGroup.top, segTop);
            borderGroup.bottom = Math.max(borderGroup.bottom, segBottom);
          } else {
            flushBorderGroup();
            borderGroup = {
              border: activeBorder,
              left: x, right: x + decoW, top: segTop, bottom: segBottom,
            };
          }
        } else {
          flushBorderGroup();
        }

        // Track-changes overlay: paint insertions / deletions in the author's
        // colour with the canonical Word markup (underline for insertions,
        // strikethrough for deletions). The author hash gives stable colours
        // for the same reviewer across pages. Disabled when
        // `showTrackChanges: false` (the "Final / No Markup" view).
        const revActive = state.showTrackChanges && !!s.revision;
        const revColor = revActive ? authorColor(s.revision!.author) : null;
        let glyphColor: string;
        // ECMA-376 §17.3.2.6 — effective background behind the glyphs, most-
        // specific first: the RUN shading (§17.3.2.32 `<w:shd>`, immediately
        // behind the glyphs — inverse-video), else the paragraph shading
        // (§17.3.1.31 `<w:pPr><w:shd>`), else the enclosing container background
        // (`state.containerShading` — the table cell fill §17.4.33, threaded by
        // renderCell). The parser filters `fill="auto"`/non-hex at every level,
        // so a non-null value here is a real paint.
        const effBg = s.background ?? para.shading ?? state.containerShading ?? null;
        if (revColor) {
          glyphColor = revColor;
        } else if (s.color) {
          glyphColor = `#${s.color}`;
        } else if (s.colorAuto || effBg != null) {
          // §17.3.2.6 (w:color) / ST_HexColorAuto §17.18.39: the automatic color
          // picks black/white for contrast against the effective background (the
          // pick is implementation-defined — delegated to core's autoContrastColor).
          // TWO states reach it:
          //   • explicit `<w:color w:val="auto"/>` (s.colorAuto — the parser's
          //     only colorAuto producer, styles.rs);
          //   • color NEVER APPLIED in the style hierarchy: §17.3.2.6 "If this
          //     element is never applied in the style hierarchy, then the
          //     characters are set to allow the consumer to automatically choose
          //     an appropriate color based on the background color behind the
          //     run's content." The parser flattens docDefaults → styles → direct
          //     rPr into the resolved `s.color`, so `s.color == null && !colorAuto`
          //     here IS exactly that never-applied state (sample-28 p.17: the
          //     `w:fill="0C0C0C"` header cells' runs carry no w:color at any
          //     level — Word paints them white).
          // The never-applied state is gated on a NON-NULL effective background:
          // with no shading anywhere the "appropriate color against the page
          // background" is the application default text color, i.e. the public
          // `defaultTextColor` render option below — rerouting it through the
          // hard black/white pick would silently break that option (and change
          // nothing for the default black). This decision is deliberately made at
          // PAINT time, not in the parser: marking every color-less run as auto
          // there would lose the resolved-vs-defaulted distinction the option
          // depends on, while the background composition only exists here.
          glyphColor = autoContrastColor(effBg);
        } else {
          glyphColor = defaultColor;
        }
        ctx.fillStyle = glyphColor;
        // Draw the glyphs. Four cases, all anchored to the WHOLE-string
        // cumulative advance so the browser's contextual CJK metrics (most
        // visibly 約物半角, the half-width collapse of （「」。）) are honoured and
        // the painted advance equals the segment's box exactly:
        //   1. §17.3.2.14 fitText: resolved per-gap, with no trailing gap after
        //      the region's last glyph and no cached w:spacing contribution.
        //   2. Character grid active on a pure-EA segment (segGridDelta !== 0):
        //      walk every glyph, advancing each to its cell start
        //      `measure(prefix) + i·Δ + justGaps·perGap`. The final glyph lands so
        //      the segment edge is measure(whole) + len·Δ + nGaps·perGap =
        //      measuredWidth + internalStretch — measure==draw by construction
        //      (§17.6.5). Folds in any justification pitch at the same time.
        //   3. Justified inter-CJK pitch only (no grid): the existing
        //      `justifiedPiecePositions` slice-at-gaps path.
        //   4. Neither: a single fillText (the common path).
        const segmentGridDeltaPx = segmentCharacterGridDeltaPx(s, drawGridDeltaPx);
        const segGridDelta = gridSegDeltaPx(drawText, segmentGridDeltaPx);
        // ECMA-376 §17.3.1.6 `<w:bidi>` (issue #929) — a segment's TRAILING
        // whitespace (an inter-word space at its logical end) must sit on the
        // segment's physical LEFT under an RTL visual frame, toward the next
        // reading word. Canvas is asked to do this via `ctx.direction='rtl'`, but
        // that is BACKEND-DEPENDENT: Chrome reorders the trailing space to the
        // left, whereas skia-canvas (the server/VRT/MCP rendering backend)
        // left-anchors the logical string and leaves the space on the physical
        // RIGHT — so the space lands on the wrong (outer) side and the word renders
        // FLUSH against its reading-next neighbour (the gap collapses; most visible
        // as a two-word label / table cell where the single inter-word gap is lost).
        // Position the whitespace EXPLICITLY instead: draw the trailing-whitespace-
        // TRIMMED glyphs (`glyphText`) shifted rightward by the whitespace advance
        // (`rtlWsShiftPx`) so the space occupies the box's LEFT — identical output
        // in both backends.
        //
        // The shift is derived from the SAME single advance authority the measure
        // pass used for the segment box (`segAdvanceWidth`: natural glyph width ×
        // §17.3.2.43 `w:w` scale + one per-code-point pitch — §17.3.2.35
        // `w:spacing`, or the §17.3.2.14 fitText per-gap), NOT by re-measuring
        // under a paint letterSpacing: the fixed pitch is per code point and does
        // NOT stretch with `w:w`, so measuring with `letterSpacing=spacing` and
        // multiplying by the scale would wrongly scale the pitch. With the
        // authority, the anchored glyphs' right edge lands exactly on the box edge
        // under every pitch combination (measure==paint).
        //
        // Consumers: the plain / §17.3.2.35 spacing / §17.3.2.43 w:w branches use
        // `glyphText`/`glyphDrawX`; the §17.3.2.14 fitText branch composes
        // `rtlWsShiftPx` with its region-end pad shift (`fitDrawX`). The docGrid
        // branch (`segGridDelta !== 0`) and the justified split-piece branch are
        // exempt: both require CJK content inside the segment (a pure-EA grid
        // segment / inter-CJK split points), which resolves to an even (LTR) bidi
        // level, so an RTL-direction segment cannot reach them outside the
        // rtl-marked EA-punctuation corner (bidi justification is already
        // approximate there — see the decoW note above). LTR segments and the
        // non-bidi fast path keep `glyphText===drawText` / `glyphDrawX===x`
        // (byte-identical). Decorations, `onTextRun`, and the pen advance stay on
        // the untrimmed box (`x` / `spanW`).
        let glyphText = drawText;
        let glyphDrawX = x;
        let rtlWsShiftPx = 0;
        if (
          visual &&
          visual.rtl[si] === true &&
          !verticalUpright &&
          /\s$/u.test(drawText)
        ) {
          const trimmed = drawText.replace(/\s+$/u, '');
          if (trimmed.length > 0) {
            // Natural (pitch-free) advances, mirroring the layout measure pass
            // (see modeledAdvance): the authority folds the pitch in itself.
            const prevLetterSpacing = ctx.letterSpacing;
            ctx.letterSpacing = '0px';
            const naturalFull = ctx.measureText(drawText).width;
            const naturalTrimmed = ctx.measureText(trimmed).width;
            ctx.letterSpacing = prevLetterSpacing;
            rtlWsShiftPx =
              segAdvanceWidth({ ...s, text: drawText }, naturalFull, drawGridDeltaPx, scale) -
              segAdvanceWidth({ ...s, text: trimmed }, naturalTrimmed, drawGridDeltaPx, scale);
            glyphText = trimmed;
            glyphDrawX = x + rtlWsShiftPx;
          }
        }
        const glyphUnitScale = useCanonicalTransform ? scale : 1;
        const glyphBaseline = useCanonicalTransform ? 0 : baseline + yOffset;
        const glyphLocalX = (absolutePaintX: number): number =>
          useCanonicalTransform ? (absolutePaintX - x) / scale : absolutePaintX;
        if (useCanonicalTransform) {
          ctx.save();
          ctx.translate(x, baseline + yOffset);
          ctx.scale(scale, scale);
        }
        if (verticalUpright && s.tateChuYoko) {
          // ECMA-376 §17.3.2.10 縦中横 (horizontal-in-vertical): draw the whole run
          // horizontally, side by side, inside ONE cell of the vertical column.
          // The cell's along-column advance is `spanW` (= s.measuredWidth, which
          // segAdvanceWidth pinned to one em for a 縦中横 seg — measure==paint).
          // `w:w` (segCharScale) compresses the digits' cross-column width;
          // vertCompress fits their height to the cell. See vertical-text.ts.
          drawTateChuYokoRun(
            ctx,
            drawText,
            x,
            baseline + yOffset,
            effSizePx,
            spanW,
            segCharScale,
            !!s.tateChuYokoCompress,
          );
        } else if (verticalUpright) {
          // ECMA-376 §17.6.20 (tbRl) — the run flows DOWN the column (logical
          // +x). Draw each glyph advancing by its measured horizontal width
          // (× the §17.3.2.43 `w:w` scale) plus the combined per-glyph pitch —
          // the docGrid cell delta (non-zero only on a pure-EA segment) plus the
          // §17.3.2.35 `w:spacing` pitch, the SAME `segLetterSpacingPx` value the
          // measured advance folds in (measure==paint) — counter-rotating upright
          // (CJK) glyphs so they stand up inside the +90°-rotated page while
          // Latin/digits stay sideways. The horizontal-only justify slicing
          // (cases below) does not apply in vertical stage-1 (the sample's
          // columns are start-aligned).
          drawVerticalRun(
            ctx,
            drawText,
            x,
            baseline + yOffset,
            effSizePx,
            segLetterSpacingPx(s, drawGridDeltaPx, scale),
            segCharScale,
            // #1014 — grow a vo=Tr rotate mark's cell to its ink ONLY when the layout
            // advance was grown by the same deficit (`s.verticalRun`), so paint==measure.
            s.verticalRun === true,
          );
        } else if (s.fitTextPerGapPx !== undefined) {
          // ECMA-376 §17.3.2.14 Manual Run Width. Same draw model as the
          // §17.18.44 FULLY-distributed arm below: the resolved region gap opens
          // at EVERY internal code-point boundary, so the whole
          // contextually-shaped string is painted in ONE fillText with a uniform
          // `ctx.letterSpacing = perGap` — glyph i lands at
          // measure(prefix_i) + i·perGap and the final glyph reaches
          // measure(whole) + (n−1)·perGap, the segment's canonical advance
          // (measure==paint; no piece slicing is needed when every boundary is a
          // gap). The canonical measuredWidth already includes one trailing
          // boundary gap on every NON-last region segment and none on the last;
          // the normal pen advance supplies that cross-segment gap. Composed
          // with §17.3.2.43 `w:w` exactly like the sibling arms: the fixed pitch
          // is divided by `segCharScale` so the ×scale frame reproduces its
          // un-scaled magnitude.
          // ECMA-376 §17.3.2.14 (Manual Run Width) + UAX#9 rule L2: mirror the
          // docx #830 RTL tab-stop leading-edge rule. A region's residual pad
          // trails its LAST glyph in READING order — the physical right under an
          // LTR base (the pen advance already leaves it there), but the physical
          // LEFT under an RTL base. So when this region-end segment draws in the
          // RTL visual frame, shift the glyph origin rightward by the pad so the
          // glyph sits at the leading (right) edge and the pad falls to its left.
          // (Non-end / multi-char segments carry trailingPad == 0 ⇒ no shift, and
          // every LTR segment keeps a zero offset ⇒ byte-identical.)
          //
          // Issue #929 composes here exactly like the sibling arms: an RTL
          // segment's TRAILING whitespace (a run-boundary space kept inside the
          // fit region) must also fall to the glyphs' LEFT, so the whitespace-
          // trimmed `glyphText` draws at `fitDrawX + rtlWsShiftPx` — the pad AND
          // the whitespace advance (its glyph width plus its per-gap share, per
          // the segAdvanceWidth authority above) are both reserved on the left,
          // and the trimmed glyphs' right edge stays on the box edge. LTR /
          // whitespace-less segments have `rtlWsShiftPx === 0` and
          // `glyphText === drawText` (byte-identical).
          const fitRtl = !!(visual && visual.rtl[si]);
          const fitPad = s.fitTextTrailingPadPx ?? 0;
          const fitDrawX = x + (fitRtl ? fitPad : 0) + rtlWsShiftPx;
          const scaled = segCharScale !== 1;
          const prevLetterSpacing = ctx.letterSpacing;
          const fitLocalX = glyphLocalX(fitDrawX);
          if (scaled) { ctx.save(); ctx.translate(fitLocalX, 0); ctx.scale(segCharScale, 1); }
          ctx.letterSpacing = `${s.fitTextPerGapPx / glyphUnitScale / segCharScale}px`;
          ctx.fillText(glyphText, scaled ? 0 : fitLocalX, glyphBaseline);
          ctx.letterSpacing = prevLetterSpacing;
          if (scaled) ctx.restore();
        } else if (segGridDelta !== 0) {
          const cps = [...drawText]; // code points (handles surrogate pairs)
          // Draw each CONTIGUOUS piece (sliced only at justify gaps) as ONE
          // contextually-shaped `fillText`, with the per-EA-glyph grid delta
          // applied via `ctx.letterSpacing`. The previous per-code-point loop
          // painted each glyph ISOLATED (no contextual shaping) yet positioned
          // glyph i by the CONTEXTUAL cumulative `measureText(prefix_i)`. JIS X
          // 4051 約物連続 packing compresses a closing-class punctuation immediately
          // followed by an opening bracket ("：［", "、［", "）（") ~half-em in
          // measureText (a bracket next to a plain kanji/kana does NOT pack), so an
          // isolated full-width bracket plus that collapsed cumulative measure
          // pulled the following glyph half-em left, OVERLAPPING the bracket.
          // Drawing the piece contiguously makes measure and draw shape the SAME
          // way (the packing honoured ⇒ no overlap), and
          // `letterSpacing = Δ` reproduces the per-cell delta the box was measured
          // with. Build `pieces` BEFORE setting letterSpacing: justifiedPiecePositions
          // is eager and its internal `measure` calls must run at the natural
          // advance (it adds `from·Δ` itself; the canvas adds Δ between glyphs
          // WITHIN each piece — together glyph i lands at measure(prefix)+i·Δ, the
          // same target as before). See @silurus/ooxml-core → justify-positions.ts.
          const measure = (str: string): number => ctx.measureText(str).width;
          // §17.3.2.35 char spacing adds to EVERY glyph (all code points), uniform
          // with the per-EA-cell grid delta on this pure-EA segment, so the two
          // combine into one per-glyph pitch. Pass the COMBINED value both as
          // `justifiedPiecePositions`' letter-spacing term (so each piece's `dx`
          // includes the accumulated pitch of the glyphs before it) and as
          // `ctx.letterSpacing` (so the canvas adds it WITHIN each piece) —
          // together glyph i lands at measure(prefix)+i·pitch (measure==paint).
          const gridPlusSpacing = segmentGridDeltaPx + segCharSpacingPx;
          // ECMA-376 §17.3.2.43 `<w:w>` (issue #816): the MEASURE pass scaled the
          // natural glyph advance by `segCharScale` (segAdvanceWidth) but left the
          // fixed per-cell pitches (grid delta, char spacing, justify slack)
          // un-scaled — w:w stretches glyphs, not the cell gaps. When a run carries
          // w:w, reproduce that at paint by drawing inside a horizontal
          // `ctx.scale(segCharScale, 1)` translated to the pen `x`: the natural
          // glyph widths (and the `measure(prefix)` prefixes inside each piece's
          // `dx`) compress with the transform, while every FIXED pitch is divided
          // by `segCharScale` so that after the ×scale it lands at its intended
          // un-scaled magnitude. `segCharScale===1` (the overwhelmingly common
          // path, and the ONLY one any current fixture hits) keeps the prior draw
          // exactly: no transform, pieces at `x + dx`, pitch un-divided.
          const scaled = segCharScale !== 1;
          const pieces = justifiedPiecePositions(
            cps,
            stretch?.splitBefore ?? [],
            distPerGap / glyphUnitScale / segCharScale,
            measure,
            gridPlusSpacing / glyphUnitScale / segCharScale,
          );
          const prevLetterSpacing = ctx.letterSpacing;
          const lineLocalX = glyphLocalX(x);
          if (scaled) { ctx.save(); ctx.translate(lineLocalX, 0); ctx.scale(segCharScale, 1); }
          const originX = scaled ? 0 : lineLocalX;
          ctx.letterSpacing = `${gridPlusSpacing / glyphUnitScale / segCharScale}px`;
          for (const { text: piece, dx } of pieces) {
            ctx.fillText(piece, originX + dx, glyphBaseline);
          }
          ctx.letterSpacing = prevLetterSpacing;
          if (scaled) ctx.restore();
        } else if (stretch && stretch.splitBefore.length > 0) {
          // ECMA-376 §17.18.44 `both`/`distribute` inter-CJK justification pitch.
          // Anchor each sliced piece to the WHOLE-string cumulative advance plus
          // the accumulated pitch, instead of summing the isolated pieces'
          // advances. That sum drifts wider than the segment's box and would paint
          // the next run over this segment's tail (most visible at a CJK→Latin
          // boundary). See `@silurus/ooxml-core` → text/justify-positions.ts.
          //
          // ECMA-376 §17.3.2.43 `<w:w>` (issue #816): the MEASURE pass scaled the
          // natural glyph advance by `segCharScale` (segAdvanceWidth) while leaving
          // the justify slack un-scaled (w:w stretches glyphs, not the distributed
          // gaps). When a run carries w:w, reproduce that by drawing inside a
          // horizontal `ctx.scale(segCharScale, 1)` translated to the pen `x`: the
          // natural glyph widths (and each piece's `measure(prefix)` prefix)
          // compress with the transform, while the justify pitch (distPerGap) and
          // §17.3.2.35 char spacing are divided by `segCharScale` so they land at
          // their intended un-scaled magnitude. `segCharScale===1` (the common
          // path, and the only one any current fixture hits) keeps the prior draw
          // exactly: no transform, pieces at `x + dx`, pitch un-divided.
          const cps = [...drawText]; // code points (handles surrogate pairs)
          const scaled = segCharScale !== 1;
          const lineLocalX = glyphLocalX(x);
          const originX = scaled ? 0 : lineLocalX;
          const prevLetterSpacing = ctx.letterSpacing;
          if (scaled) { ctx.save(); ctx.translate(lineLocalX, 0); ctx.scale(segCharScale, 1); }
          if (stretch.splitBefore.length === cps.length - 1) {
            // FULLY distributed: a gap was opened at EVERY inter-glyph boundary
            // (pure-CJK justify), so the pitch is UNIFORM. Drawing one glyph per
            // piece in isolation (the loop below) loses the browser's contextual
            // packing — JIS X 4051 約物連続 compresses a closing-class punctuation
            // immediately followed by an opening bracket ("：［", "、［", "）（")
            // ~half-em in measureText — so the bracket paints full-width while the
            // next glyph, positioned by the COLLAPSED cumulative measure, is pulled
            // left and OVERLAPS it (the justify analog of the docGrid case-1 fix,
            // PR #626).
            // Draw the whole CONTEXTUALLY-shaped run in ONE fillText with
            // ctx.letterSpacing = distPerGap: glyph i lands at measure(prefix_i) +
            // i·distPerGap (the exact justified position), so the final glyph
            // reaches measure(whole) + (n-1)·distPerGap = the segment box edge
            // (= internalStretch). Restore the prior letterSpacing afterwards; no
            // measureText runs inside the set/restore window.
            // §17.3.2.35 char spacing is a per-glyph pitch on top of the justify
            // slack; both add uniformly, so combine them (the box measured
            // len·charSpacingPx separately from the justify slack — measure==paint).
            // Both fixed pitches are divided by `segCharScale` (see the arm header)
            // so the ×scale frame reproduces their un-scaled magnitude (a no-op
            // divide by 1 on the common non-w:w path).
            ctx.letterSpacing = `${(distPerGap + segCharSpacingPx) / glyphUnitScale / segCharScale}px`;
            ctx.fillText(drawText, originX, glyphBaseline);
          } else {
            const measure = (str: string): number => ctx.measureText(str).width;
            // Partial justify split: pass the char-spacing pitch both as the
            // per-glyph letter-spacing term (so each piece's `dx` includes the
            // spacing of the glyphs before it) and as `ctx.letterSpacing` (so it
            // is added WITHIN each piece) — measure==paint across the split. Both
            // the justify slack and the char spacing are divided by `segCharScale`
            // so they survive the ×scale frame un-stretched (see the arm header).
            for (const { text: piece, dx } of justifiedPiecePositions(
              cps,
              stretch.splitBefore,
              distPerGap / glyphUnitScale / segCharScale,
              measure,
              segCharSpacingPx / glyphUnitScale / segCharScale,
            )) {
              ctx.letterSpacing = `${segCharSpacingPx / glyphUnitScale / segCharScale}px`;
              ctx.fillText(piece, originX + dx, glyphBaseline);
            }
          }
          ctx.letterSpacing = prevLetterSpacing;
          if (scaled) ctx.restore();
        } else if (segCharScale !== 1) {
          // ECMA-376 §17.3.2.43 `<w:w>` — draw each glyph at `segCharScale`× its
          // normal width. Canvas has no per-glyph width scale, so paint under a
          // horizontal `ctx.scale`: translate to the run's pen x, scale x only,
          // and draw at local origin. Char spacing (if any) is applied in the
          // UNSCALED point space, so set `letterSpacing = charSpacing / scale`
          // inside the scaled frame to keep the fixed pitch un-stretched by w:w.
          // (The docGrid and justify arms above compose the SAME transform when a
          // grid / distributed run also carries w:w — issue #816.)
          ctx.save();
          ctx.translate(glyphLocalX(glyphDrawX), 0);
          ctx.scale(segCharScale, 1);
          const prevLetterSpacing = ctx.letterSpacing;
          if (segCharSpacingPx !== 0) {
            ctx.letterSpacing = `${segCharSpacingPx / glyphUnitScale / segCharScale}px`;
          }
          ctx.fillText(glyphText, 0, glyphBaseline);
          ctx.letterSpacing = prevLetterSpacing;
          ctx.restore();
        } else if (segCharSpacingPx !== 0) {
          // §17.3.2.35 `<w:spacing>` only (no grid, no justify, no scale): the
          // whole run draws with a uniform per-glyph letter-spacing pitch that the
          // layout already folded into `s.measuredWidth` (measure==paint).
          const prevLetterSpacing = ctx.letterSpacing;
          ctx.letterSpacing = `${segCharSpacingPx / glyphUnitScale}px`;
          ctx.fillText(glyphText, glyphLocalX(glyphDrawX), glyphBaseline);
          ctx.letterSpacing = prevLetterSpacing;
        } else {
          ctx.fillText(glyphText, glyphLocalX(glyphDrawX), glyphBaseline);
        }
        if (useCanonicalTransform) {
          ctx.restore();
          // Overlay consumers receive paint-space geometry and must see the
          // corresponding paint-size font even though the glyphs themselves were
          // shaped at scale 1 and mapped through the local viewport transform.
          ctx.font = buildFont(s.bold, s.italic, effSizePx, s.fontFamily, fontFamilyClasses, s.fontRoute);
        }
        // §17.3.2.19 — restore the inherited font-kerning now the run's glyphs are
        // painted (the following ruby / emphasis-mark draws are separate glyphs at
        // their own sizes and use the inherited kerning). No-op when unset.
        if (s.kerning != null) ctx.fontKerning = prevFontKerning;

        // Ruby annotation: small text centered above the base glyphs.
        if (s.ruby) {
          const rubySizePx = s.ruby.fontSizePt * scale;
          const rubyFont = buildFont(s.bold, s.italic, rubySizePx, s.fontFamily, fontFamilyClasses, s.fontRoute);
          ctx.save();
          ctx.font = rubyFont;
          const rubyW = ctx.measureText(s.ruby.text).width;
          const rubyX = x + (spanW - rubyW) / 2;
          // Sit the ruby's baseline a small gap above the base ascent so the
          // characters don't touch. fillText baseline is at the line of the
          // characters, so subtract the ruby descent + small gap from the
          // base's ascent line to position correctly.
          const rubyBaseline = s.ruby.hpsRaisePt != null
            ? baseline + yOffset - s.ruby.hpsRaisePt * scale
            : baseline + yOffset - effSizePx * 0.85 - rubySizePx * 0.1;
          // Ruby shares glyphColor, so a track-changes run's ruby inherits the
          // author revision color (a behavior change from previously ignoring
          // revColor for ruby).
          ctx.fillStyle = glyphColor;
          ctx.fillText(s.ruby.text, rubyX, rubyBaseline);
          ctx.restore();
        }

        // ECMA-376 §17.3.2.12 emphasis mark (圏点): a small glyph stamped on every
        // NON-SPACE character (§17.18.24), centred above each glyph (below for
        // `underDot`). Drawn AFTER the text so it overlays; the advance is
        // unchanged (no layout impact). The per-glyph centre uses the SAME
        // contextual `measureText` cumulative advance the glyph draw is anchored
        // to, plus the run's uniform per-glyph pitch so a docGrid cell delta /
        // fully-distributed justify pitch keeps the mark centred. (A partial
        // justify split — non-uniform pitch — falls back to pitch 0, which
        // stays within a fraction of a glyph of centre and is not worth the
        // complexity of re-deriving the sliced positions.)
        if (s.emphasisMark) {
          const geom = emphasisMarkGeometry(s.emphasisMark, effSizePx);
          // Uniform per-glyph pitch matching the case the glyphs were drawn with.
          const fullyDistributed =
            !!stretch &&
            stretch.splitBefore.length > 0 &&
            stretch.splitBefore.length === [...s.text].length - 1;
          const markPitch =
            segGridDelta !== 0
              ? segmentGridDeltaPx
              : fullyDistributed
                ? distPerGap
                : 0;
          const measureMark = (str: string): number => ctx.measureText(str).width;
          const centers = emphasisMarkCenters(s.text, measureMark, x, markPitch);
          // Above marks sit a small gap above the glyph box top (the same
          // ~0.85em ascent the box decorations use); below marks (underDot) sit
          // just under the box bottom (baseline + ~0.25em). The gap keeps the
          // mark clear of the glyph without stealing line height.
          const markGap = effSizePx * 0.06;
          const markCy = geom.above
            ? boxTop - markGap - geom.radius
            : boxTop + boxHeight + markGap + geom.radius;
          ctx.save();
          ctx.fillStyle = glyphColor;
          ctx.strokeStyle = glyphColor;
          for (const { centerX } of centers) {
            if (geom.shape === 'circle') {
              // Hollow circle (§17.18.24 "circle"): stroked ring.
              ctx.lineWidth = Math.max(0.5, geom.radius * 0.35);
              ctx.beginPath();
              ctx.arc(centerX, markCy, geom.radius, 0, Math.PI * 2);
              ctx.stroke();
            } else if (geom.shape === 'comma') {
              // Sesame / comma mark (§17.18.24 "comma"): a filled teardrop —
              // a disc with a short tail down-right, approximating the boten
              // sesame «﹅». Kept simple (disc + triangle) so it reads at body
              // sizes without a font dependency.
              ctx.beginPath();
              ctx.arc(centerX, markCy, geom.radius, 0, Math.PI * 2);
              ctx.fill();
              ctx.beginPath();
              ctx.moveTo(centerX - geom.radius * 0.5, markCy + geom.radius * 0.2);
              ctx.lineTo(centerX + geom.radius * 0.5, markCy + geom.radius * 0.2);
              ctx.lineTo(centerX - geom.radius * 0.1, markCy + geom.radius * 1.4);
              ctx.closePath();
              ctx.fill();
            } else {
              // Filled dot (§17.18.24 "dot" / "underDot").
              ctx.beginPath();
              ctx.arc(centerX, markCy, geom.radius, 0, Math.PI * 2);
              ctx.fill();
            }
          }
          ctx.restore();
        }

        if (state.onTextRun && s.text) {
          // ECMA-376 §17.6.20 (tbRl) — `x`/`state.y` are LOGICAL flow coords; on a
          // vertical page the overlay DOM lives on the physical (rotated) canvas,
          // so project the logical top-left to physical and hand the span the +90°
          // rotation. `verticalTextLayerPlacement` returns null on horizontal pages
          // (the span stays at the logical `x`/`y`, byte-identical to before).
          const place = verticalTextLayerPlacement(
            x, state.y, state.verticalPhys?.cssWidthPx ?? 0, !!state.verticalCJK,
          );
          // Reuse the paint path's single pitch authority so selection and find
          // overlays reproduce §17.3.2.14 fitText or docGrid + §17.3.2.35 spacing.
          // Upright-vertical / 縦中横 runs retain their existing payload and
          // geometry; an all-rotated (btLr) run drew through the horizontal
          // branch WITH the pitch, so its overlay reports it like horizontal.
          const letterSpacingPx = !verticalUpright && !s.tateChuYoko
            ? segLetterSpacingPx(s, drawGridDeltaPx, scale)
            : 0;
          state.onTextRun(textRunPaintInfo({
            text: s.text,
            x: place ? place.left : x,
            y: place ? place.top : state.y,
            w: spanW,
            h: lineH,
            fontSize: effSizePx,
            font: ctx.font,
            ...(letterSpacingPx !== 0 ? { letterSpacingPx } : {}),
            transform: place?.transform,
            // IX1 — hand the resolved hyperlink target to the overlay so a link
            // run becomes clickable. Undefined for non-link runs (no payload
            // change). Does not touch any drawing above.
            hyperlink: s.hyperlink,
            // §17.3.2.10 縦中横 — flag a tate-chu-yoko run so the overlays clamp
            // their extent to the drawn one-em cell (`w`) instead of the run's
            // natural glyph width (#836). Only set on a vertical page, where the
            // 縦中横 draw path (above) actually fires; `undefined` otherwise so a
            // non-縦中横 run's payload is byte-identical.
            eastAsianVert: verticalUpright && s.tateChuYoko ? true : undefined,
          }));
        }

        // Underline / strike share the glyph colour, so an inverse-video run
        // (automatic colour on a dark background) draws a white rule too.
        const lineColor = glyphColor;
        const lineW = Math.max(0.5, effSizePx * 0.05);
        // Crispness nudge (see crispOffset): underline / strike-through are
        // horizontal strokes; each snaps onto the nearest crisp device row from
        // its own y (an odd device-width one would otherwise straddle two rows).
        // Compute the offset per line because each stroke sits at a different y.
        // Underline / strike run the SAME `decoW` as the box decorations: the
        // segment's grid-aware advance (§17.6.5) plus the interior + owned
        // trailing-space justification pitch. Word runs the rule under a run's
        // spaces (incl. their justified widening), so the line decoration tracks
        // the drawn advance and stays flush with the box fills (one width concept
        // for every decoration, matching the pptx renderer's run rules).
        const isInsertion = revActive && s.revision?.kind === 'insertion';
        const isDeletion = revActive && s.revision?.kind === 'deletion';

        if (s.underline || isInsertion) {
          // The docx underline anchor (byte-stable across releases): the single
          // rule sits `effSizePx*0.12` below the baseline, at weight `lineW`.
          const uyRaw = baseline + yOffset + effSizePx * 0.12;
          // A styled underline (§17.3.2.40 `<w:u w:val>` other than single) is
          // drawn by the shared core painter (§20.1.10.82 ST_TextUnderlineType).
          // Insertions carry no style, so they always take the single path.
          // `s.underlineColor` (§17.3.2.40 `w:u@color`) overrides the glyph
          // colour when a concrete hex is given; `auto` (or absent) follows the
          // glyph colour.
          const uStyle = !isInsertion ? s.underlineStyle : undefined;
          if (uStyle) {
            const uColor =
              s.underlineColor && s.underlineColor !== 'auto'
                ? `#${s.underlineColor}`
                : lineColor;
            // core.drawUnderline computes its own rule y as `baseline +
            // max(2, coreLineW)`. Pass a `baseline` shifted so that lands exactly
            // on the docx anchor `uyRaw`, keeping styled underlines flush with the
            // single rule's position. coreLineW mirrors core's own weight formula.
            const coreLineW = Math.max(1, effSizePx * 0.05);
            const coreBaseline = uyRaw - Math.max(2, coreLineW);
            drawUnderline(
              ctx,
              x,
              coreBaseline,
              decoW,
              effSizePx,
              uColor,
              docxUnderlineToDrawingML(uStyle),
              state.dpr,
            );
            ctx.setLineDash([]);
          } else {
            const uColor =
              !isInsertion && s.underlineColor && s.underlineColor !== 'auto'
                ? `#${s.underlineColor}`
                : lineColor;
            ctx.strokeStyle = uColor;
            ctx.lineWidth = lineW;
            const uy = uyRaw + crispOffset(uyRaw, lineW, state.dpr);
            ctx.beginPath(); ctx.moveTo(x, uy); ctx.lineTo(x + decoW, uy); ctx.stroke();
          }
        }

        if (s.strikethrough || isDeletion) {
          ctx.strokeStyle = lineColor;
          ctx.lineWidth = lineW;
          const syRaw = baseline + yOffset - effSizePx * 0.3;
          const sy = syRaw + crispOffset(syRaw, lineW, state.dpr);
          ctx.beginPath(); ctx.moveTo(x, sy); ctx.lineTo(x + decoW, sy); ctx.stroke();
        }

        if (s.doubleStrikethrough) {
          ctx.strokeStyle = lineColor;
          ctx.lineWidth = lineW;
          const sy1Raw = baseline + yOffset - effSizePx * 0.35;
          const sy2Raw = baseline + yOffset - effSizePx * 0.22;
          const sy1 = sy1Raw + crispOffset(sy1Raw, lineW, state.dpr);
          const sy2 = sy2Raw + crispOffset(sy2Raw, lineW, state.dpr);
          ctx.beginPath(); ctx.moveTo(x, sy1); ctx.lineTo(x + decoW, sy1); ctx.stroke();
          ctx.beginPath(); ctx.moveTo(x, sy2); ctx.lineTo(x + decoW, sy2); ctx.stroke();
        }
      }

      // Advance the pen past the segment's glyphs plus any internal justification
      // pitch added between them.
      x += s.measuredWidth + internalStretch;
      // Trailing inter-segment gap (an inter-word space or inter-CJK boundary at
      // this segment's edge), applied AFTER the segment so the next one starts
      // shifted. distributeLineSlack only sets trailingGap on gap-opening
      // segments — never the visually-final segment or a leading-indent segment —
      // so the final glyph still lands on the margin (Σgaps == slack).
      if (trailingDistributionGap) x += distPerGap;
    }
    // End of line closes any open run-border group: a frame never spans lines
    // (each line wrap starts a fresh box on the next line).
    flushBorderGroup();
    if (paraNeedsBidi) ctx.direction = 'ltr'; // reset for subsequent draws

    state.y += lineH;
}

// ===== Text layout =====

/** Fill a tab gap with its leader characters (e.g. TOC dot leaders, ECMA-376 §17.3.1.37). */
function drawTabLeader(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  leader: NonNullable<LayoutTabSeg['leader']>,
  x: number,
  baseline: number,
  width: number,
  fontPx: number,
  color: string,
  bold?: boolean,
  italic?: boolean,
): void {
  const ch =
    leader === 'hyphen'
      ? '-'
      : leader === 'underscore' || leader === 'heavy'
        ? '_'
        : leader === 'middleDot'
          ? '·'
          : '.';
  ctx.save();
  // ECMA-376 §17.3.1.37: the leader fill takes the formatting of the tab's run,
  // so a bold/italic TOC entry draws a bold/italic dot leader.
  const style = `${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}`;
  ctx.font = `${style}${fontPx}px serif`;
  ctx.fillStyle = color;
  const chW = ctx.measureText(ch).width;
  if (chW > 0) {
    // Dots sit on a loose grid; other leaders are drawn solid.
    const step = leader === 'dot' || leader === 'middleDot' ? chW * 1.5 : chW;
    const margin = chW * 0.5;
    // Leave a clear gap (about one dot-step) before the page number so the
    // leader doesn't run right up against it.
    const end = x + width - step - margin;
    for (let cx = x + margin; cx <= end; cx += step) {
      ctx.fillText(ch, cx, baseline);
    }
  }
  ctx.restore();
}

function renderInlineImage(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  seg: LayoutImageSeg,
  x: number,
  baseline: number,
  scale: number,
  images: Map<string, DecodedImage>,
  vertical: boolean,
): void {
  // Anchor images are skipped during layout (measuredWidth=0, not added to line.segments)
  // and are drawn later by renderAnchorImages — so this function only handles inline images.
  if (seg.anchor) return;
  const w = seg.widthPt * scale;
  const h = seg.heightPt * scale;
  const boxY = baseline - h;
  // ECMA-376 §17.6.20 (tbRl) — an inline image/chart is a graphic, not text, so it
  // stands UPRIGHT inside the +90°-rotated page. `drawUprightBox` counter-rotates
  // the flow box (logical `x,boxY,w,h`) −90° about its centre and invokes the
  // callback with the un-swapped upright rect, so the image/chart is painted right
  // way up. On horizontal pages the callback runs with the box unchanged (no
  // rotation), byte-identical to the pre-vertical inline draw.
  const paint = (
    draw: (dx: number, dy: number, dw: number, dh: number) => void,
  ): void => {
    if (vertical) drawUprightBox(ctx, x, boxY, w, h, draw);
    else draw(x, boxY, w, h);
  };
  // ECMA-376 §21.2 inline chart: paint through the shared core chart renderer
  // (the same entry point pptx/xlsx use), at the inline box's top-left. `scale`
  // is px-per-pt in this renderer, which is exactly the `ptToPx` renderChart
  // wants to scale the chart's point-sized fonts/axes — so pass it straight.
  if (seg.chart) {
    const chart = seg.chart;
    paint((dx, dy, dw, dh) =>
      renderChart(ctx as CanvasRenderingContext2D, chart, { x: dx, y: dy, w: dw, h: dh }, scale),
    );
    return;
  }
  const bmp = images.get(imageKey(seg.imagePath, seg.colorReplaceFrom, seg.duotone));
  if (!bmp) return;
  // §20.1.8.6 alphaModFix — multiply the picture's opacity for the draw.
  const hasAlpha = seg.alpha != null && seg.alpha < 1;
  if (hasAlpha) {
    ctx.save();
    ctx.globalAlpha *= seg.alpha as number;
  }
  paint((dx, dy, dw, dh) => drawImageRunBitmap(ctx, bmp, seg, dx, dy, dw, dh));
  if (hasAlpha) ctx.restore();
}

/** Paint a picture with the effective DrawingML rotation/reflection composed
 * by the parser according to Annex L §L.4.7.4–§L.4.7.6. */
function drawImageRunBitmap(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  bmp: DecodedImage,
  image: Pick<ImageRun, 'srcRect' | 'rotation' | 'flipH' | 'flipV'>,
  x: number,
  y: number,
  w: number,
  h: number,
): void {
  const rotation = image.rotation ?? 0;
  if (rotation === 0 && !image.flipH && !image.flipV) {
    drawImageCropped(ctx, bmp, image.srcRect ?? undefined, x, y, w, h);
    return;
  }
  ctx.save();
  ctx.translate(x + w / 2, y + h / 2);
  ctx.rotate((rotation * Math.PI) / 180);
  ctx.scale(image.flipH ? -1 : 1, image.flipV ? -1 : 1);
  drawImageCropped(ctx, bmp, image.srcRect ?? undefined, -w / 2, -h / 2, w, h);
  ctx.restore();
}

/** Collect and draw anchor images with wrapMode='none' (or unspecified).
 * Wrap floats (square/topAndBottom/tight/through) are drawn by registerAnchorFloats.
 *
 * `phase` = 'behind' draws only shapes with behindDoc=true (sorted by zOrder asc);
 * `phase` = 'front' draws shapes without behindDoc + all anchor images. Front
 * shapes are sorted by `wp:anchor/@relativeHeight` (lower first, higher on top)
 * while non-shape anchors keep their legacy run-order fallback. */
function renderAnchorImages(
  para: DocParagraph,
  state: RenderState,
  paragraphTopPx: number,
  phase: 'behind' | 'front' = 'front',
  wrapFloatParagraphTopPx = paragraphTopPx,
): void {
  if (state.dryRun) return;
  if (phase === 'behind') {
    const shapes = para.runs
      .filter((r): r is ShapeRun & { type: 'shape' } =>
        r.type === 'shape' && !!(r as unknown as ShapeRun).behindDoc)
      .slice()
      .sort((a, b) =>
        ((a as unknown as ShapeRun).zOrder ?? 0) - ((b as unknown as ShapeRun).zOrder ?? 0));
    for (const s of shapes) {
      const shape = s as unknown as ShapeRun;
      const top = isWrapFloat(shape.wrapMode) ? wrapFloatParagraphTopPx : paragraphTopPx;
      renderAnchorShape(shape, state, top);
    }
    return;
  }
  const frontRuns = para.runs
    .map((run, index) => {
      const shapeZ = run.type === 'shape'
        ? (run as unknown as ShapeRun).zOrder
        : null;
      return {
        run,
        index,
        z: typeof shapeZ === 'number' && Number.isFinite(shapeZ) ? shapeZ : index,
      };
    })
    .sort((a, b) => a.z - b.z || a.index - b.index);
  for (const { run } of frontRuns) {
    if (run.type === 'shape') {
      const s = run as unknown as ShapeRun;
      if (s.behindDoc) continue;
      const top = isWrapFloat(s.wrapMode) ? wrapFloatParagraphTopPx : paragraphTopPx;
      renderAnchorShape(s, state, top);
      continue;
    }
    if (run.type === 'chart') {
      // ECMA-376 §20.4.2.3 (`<wp:anchor>`) + §21.2 (chart) — wrap-mode
      // charts are painted by registerChartFloat after their exclusion rect is
      // registered. This branch paints only wrapNone/no-wrap anchors through
      // the same box resolver and core chart renderer.
      const chartRun = run as unknown as ChartRun & { type: 'chart' };
      if (!chartRun.anchor) continue;
      if (isWrapFloat(chartRun.wrapMode)) continue;
      const { x: pageX, y: pageY, w, h } = resolveAnchorBox(chartRun, state, paragraphTopPx);
      const chart = chartRun.chart;
      if (state.verticalCJK) {
        // §17.6.20 (tbRl) — a chart is a graphic, not text: keep it upright.
        drawUprightBox(state.ctx, pageX, pageY, w, h, (dx, dy, dw, dh) =>
          renderChart(state.ctx as CanvasRenderingContext2D, chart, { x: dx, y: dy, w: dw, h: dh }, state.scale),
        );
      } else {
        renderChart(state.ctx as CanvasRenderingContext2D, chart, { x: pageX, y: pageY, w, h }, state.scale);
      }
      continue;
    }
    if (run.type !== 'image') continue;
    const img = run as unknown as ImageRun;
    if (!img.anchor) continue;
    if (isWrapFloat(img.wrapMode)) continue;  // drawn as a float
    const bmp = state.images.get(imageKey(img.imagePath, img.colorReplaceFrom, img.duotone));
    if (!bmp) continue;

    // wrapNone images anchor against the paragraph's pre-spaceBefore top
    // (paragraphTopPx). Shared box resolution with the float path. By design the
    // box-resolution is symmetric but the overlap handling is NOT: wrap floats
    // (registerAnchorFloats) build an exclusion rect and run resolveFloatOverlap,
    // whereas wrapNone images carry no exclusion rect — they are positioned
    // directly in the paragraph flow (ECMA-376 wrapNone, §20.4.2.x: the object
    // does not displace text and is not displaced by other floats), so dist* is
    // unused here.
    const { x: pageX, y: pageY, w, h } = resolveAnchorBox(img, state, paragraphTopPx);
    // §20.1.8.6 alphaModFix — multiply the picture's opacity.
    const hasAlpha = img.alpha != null && img.alpha < 1;
    if (hasAlpha) {
      state.ctx.save();
      state.ctx.globalAlpha *= img.alpha as number;
    }
    if (state.verticalCJK) {
      // §17.6.20 (tbRl) — an anchored image is not text: keep it UPRIGHT inside
      // the +90°-rotated page by counter-rotating about its box centre.
      drawUprightBox(state.ctx, pageX, pageY, w, h, (dx, dy, dw, dh) =>
        drawImageRunBitmap(state.ctx, bmp, img, dx, dy, dw, dh),
      );
    } else {
      drawImageRunBitmap(state.ctx, bmp, img, pageX, pageY, w, h);
    }
    if (hasAlpha) state.ctx.restore();
  }
}

// Anchor placement geometry (xContainer / yContainer / resolveAnchorX /
// resolveAnchorY, ECMA-376 §20.4.3.x) lives in anchor-geometry.ts and is
// imported above; the shape/image render paths below consume it.

/** Convert a parsed docx LineEnd into core's ArrowEnd. Returns undefined when
 *  absent so the Stroke field stays unset. */
function lineEndToArrowEnd(
  end: ShapeRun['headEnd'],
): ArrowEnd | undefined {
  if (!end) return undefined;
  return { type: end.type, w: end.w, len: end.len };
}

/**
 * Resolve an anchored shape's page-space bounding box {x,y,w,h} (px). Shared by
 * renderAnchorShape (where the shape is drawn) and registerAnchorFloats (where
 * its float-exclusion rect is built), so the exclusion band matches the paint
 * box exactly — see root CLAUDE.md (no duplicated geometry).
 *
 * Mirrors the renderer's sizing: sizeRelH/sizeRelV (ECMA-376 §20.4.2.18)
 * override the static extent, and a wgp child scales by the group ratio with its
 * within-group offset scaled in step; resolveAnchorX/Y then place the box. `w`/`h`
 * may be 0/negative for degenerate line presets — the caller decides how to
 * treat those (renderAnchorShape draws a line; a wrap-shape with no area
 * registers no float).
 */
function resolveShapeBox(
  shape: ShapeRun,
  state: RenderState,
  paragraphTopPx: number,
): { x: number; y: number; w: number; h: number } {
  // ECMA-376 §17.6.20 + §20.4.3.x (issue #988 batch-3 adjudication ②): on a
  // vertical (tbRl) page an anchored shape's positionH/V resolve against the
  // PHYSICAL (un-rotated) page — the drawing layer is independent of the
  // section text direction, exactly like the image path (resolveAnchorBox).
  // Resolve in the physical frame, then project into the swapped logical
  // layout frame (w↔h swapped) so the float-exclusion band and the flow all
  // share one geometry. A `paragraph`/`line`-relative positionV anchors from
  // the PHYSICAL TOP of the anchor paragraph's COLUMN (Word GT: margin-top +
  // posOffset for a single-column body) — that physical y is the column
  // band's logical x start (`state.contentX`, since physical y = logical x
  // under the +90° page paint), NOT the paragraph's logical flow
  // `paragraphTopPx`, which lies on the column-progression axis.
  if (state.verticalPhys) {
    const phys = resolveShapeBox(
      shape,
      verticalPhysicalContentState(state),
      state.contentX,
    );
    return physicalToLogicalAnchorBox(
      phys.x, phys.y, phys.w, phys.h, state.verticalPhys.cssWidthPx,
    );
  }
  const { scale } = state;
  // ECMA-376 §20.4.2.18: when wp14:sizeRelH/sizeRelV is present it overrides
  // the static wp:extent for that axis. The size is `relativeFrom` container
  // size × pct.
  //
  // For a wgp group with sizeRelH, the parent group resizes and every child
  // shape scales proportionally — so a grouped child's effective width is
  // `original_width × (new_group_w / old_group_w)`, and its within-group
  // offset (carried by anchorXPt) scales by the same ratio. Standalone
  // shapes simply take `container × pct` as their width.
  let w = shape.widthPt * scale;
  let h = shape.heightPt * scale;
  let offsetXPt = shape.anchorXPt;
  let offsetYPt = shape.anchorYPt;
  let alignWidthPt = shape.groupWidthPt ?? null;
  let alignHeightPt = shape.groupHeightPt ?? null;
  if (shape.widthPct != null) {
    const c = xContainer(shape.widthRelativeFrom, false, state);
    const newSizePt = ((c.end - c.start) * shape.widthPct) / scale;
    if (shape.groupWidthPt != null && shape.groupWidthPt > 0) {
      const ratio = newSizePt / shape.groupWidthPt;
      w = shape.widthPt * scale * ratio;
      offsetXPt = shape.anchorXPt * ratio;
    } else {
      w = newSizePt * scale;
    }
    alignWidthPt = newSizePt;
  }
  if (shape.heightPct != null) {
    const c = yContainer(shape.heightRelativeFrom, false, paragraphTopPx, state);
    const newSizePt = ((c.end - c.start) * shape.heightPct) / scale;
    if (shape.groupHeightPt != null && shape.groupHeightPt > 0) {
      const ratio = newSizePt / shape.groupHeightPt;
      h = shape.heightPt * scale * ratio;
      offsetYPt = shape.anchorYPt * ratio;
    } else {
      h = newSizePt * scale;
    }
    alignHeightPt = newSizePt;
  }
  const x = resolveAnchorX(
    shape.anchorXAlign, shape.anchorXFromMargin, offsetXPt, w, state,
    shape.anchorXRelativeFrom, shape.pctPosH, alignWidthPt,
  );
  const y = resolveAnchorY(
    shape.anchorYAlign, shape.anchorYFromPara, offsetYPt, h, paragraphTopPx, state,
    shape.anchorYRelativeFrom, shape.pctPosV, alignHeightPt,
  );
  return { x, y, w, h };
}

/** The solid fill colour of a shape as a CSS `#rrggbb` string, or `null` when
 *  the shape has no solid fill (gradient / none). Used for watermark text. */
function shapeFillColor(fill: ShapeFill | null | undefined): string | null {
  if (fill && fill.fillType === 'solid') return `#${fill.color}`;
  return null;
}

/**
 * Draw a VML `<v:textpath>` text watermark (ECMA-376 Part 4 §19.1.2.23) into the
 * box `(x, y, w, h)` (device px). Word emits watermarks with the WordArt
 * `#_x0000_t136` shapetype, whose `fitshape` default STRETCHES the text to the
 * edges of the shape box — so the drawn size is derived from the box geometry,
 * not the nominal `font-size` in the textpath style (which Word writes as a
 * placeholder `1pt`). The text is:
 *   - measured once at a reference size to get its natural advance/height,
 *   - non-uniformly scaled so it exactly fills `w × h` (fitshape),
 *   - rotated by `rotationDeg` (clockwise, §19.1.2.19) about the box centre,
 *   - filled with `color` at `opacity` alpha (§19.1.2.5 `<v:fill opacity>`).
 *
 * The transform is applied about the box centre and the text is drawn centred
 * (`textAlign`/`textBaseline` = middle), so the watermark sits centred in its
 * box regardless of rotation. Exported for unit testing the geometry.
 */
export function drawWatermarkTextPath(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  textPath: TextPath,
  x: number,
  y: number,
  w: number,
  h: number,
  rotationDeg: number,
  color: string | null,
  opacity: number,
  fontFamilyClasses: Record<string, string> = {},
): void {
  const text = textPath.string;
  if (!text || w <= 0 || h <= 0) return;

  // Reference measurement at a fixed pixel size; the fitshape scale maps the
  // natural text box onto the shape box. REF is arbitrary (cancels out in the
  // scale ratio) but large enough to keep measureText precise.
  const REF = 100;
  ctx.save();
  ctx.font = buildFont(textPath.bold ?? false, textPath.italic ?? false, REF, textPath.fontFamily ?? null, fontFamilyClasses);
  const m = ctx.measureText(text);
  const natW = m.width || REF;
  // Prefer the font bounding box (cap-to-descender) for the natural height; fall
  // back to the em size when the platform doesn't report it.
  const asc = m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent ?? REF * 0.8;
  const desc = m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent ?? REF * 0.2;
  const natH = asc + desc || REF;

  const cx = x + w / 2;
  const cy = y + h / 2;
  ctx.translate(cx, cy);
  if (rotationDeg !== 0) ctx.rotate((rotationDeg * Math.PI) / 180);
  // fitshape: stretch the text to the box edges (non-uniform). The reference
  // font renders at REF px; scaling the axes by (w/natW, h/natH) lands the ink
  // exactly on the box.
  ctx.scale(w / natW, h / natH);
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.globalAlpha = Math.max(0, Math.min(1, opacity));
  ctx.fillStyle = color ?? '#c0c0c0';
  ctx.fillText(text, 0, 0);
  ctx.restore();
}

function renderAnchorShape(shape: ShapeRun, state: RenderState, paragraphTopPx: number): void {
  // ECMA-376 §17.6.20 + §20.4.3.x (issue #988 batch-3 adjudication ②): on a
  // vertical (tbRl) page an anchored shape stays UPRIGHT at its physical-page
  // position — the drawing layer does not rotate with the text flow, and a
  // horizontal (`vert="horz"`) text body keeps horizontal labels. Undo the +90°
  // page paint (its exact inverse: rotate −90° then translate(−cssW, 0), the
  // same physical frame the header/footer path re-enters) and re-run this
  // renderer with the PHYSICAL state view: `resolveShapeBox` then resolves the
  // physical box directly, text draws horizontally (verticalCJK off), and the
  // overlay geometry is emitted unrotated at physical coordinates. The physical
  // paragraph base for a `paragraph`-relative positionV is the column's
  // physical top = the column band's logical x (`state.contentX`), matching
  // resolveShapeBox's logical-projection branch so the float band (registered
  // from the LOGICAL projection) and the painted shape share one geometry.
  const verticalFrame = state.verticalPhys;
  if (verticalFrame) {
    const cssW = verticalFrame.cssWidthPx;
    const { ctx } = state;
    ctx.save();
    ctx.rotate(-Math.PI / 2);
    ctx.translate(-cssW, 0);
    renderAnchorShape(shape, verticalPhysicalContentState(state), state.contentX);
    ctx.restore();
    return;
  }
  const { ctx, scale } = state;
  let { x, y, w, h } = resolveShapeBox(shape, state, paragraphTopPx);
  // Line/connector presets (ECMA-376 §20.1.9.18) are valid with a degenerate
  // bounding box — a horizontal line has h==0, a vertical line w==0. Stroking
  // such a path still draws a visible segment, so only bail when there is truly
  // nothing to draw (both dimensions zero) or an inverted box (negative).
  const preset = shape.presetGeometry?.toLowerCase() ?? '';
  const isLineGeom =
    preset === 'line' ||
    preset.startsWith('straightconnector') ||
    preset.startsWith('bentconnector') ||
    preset.startsWith('curvedconnector');
  const isCalloutGeom =
    preset === 'callout1' ||
    preset === 'callout2' ||
    preset === 'callout3' ||
    preset === 'bordercallout1' ||
    preset === 'bordercallout2' ||
    preset === 'bordercallout3' ||
    preset === 'accentcallout1' ||
    preset === 'accentcallout2' ||
    preset === 'accentcallout3' ||
    preset === 'accentbordercallout1' ||
    preset === 'accentbordercallout2' ||
    preset === 'accentbordercallout3';
  // Straight / bent connectors whose leader we re-stroke retracted from filled
  // line-end decorations (so the line stops at the arrow base). Callout leader
  // lines are emitted by the preset engine as their trailing path, so they can
  // use the same retract/re-stroke path. Curved connectors are excluded — their
  // Bézier leader can't be retracted from a polyline vertex without
  // straightening it, so they keep the preset leader.
  const isRetractableLeader =
    isCalloutGeom ||
    preset === 'line' ||
    preset.startsWith('straightconnector') ||
    preset.startsWith('bentconnector');
  if (w < 0 || h < 0) return;
  if (isLineGeom ? w === 0 && h === 0 : w === 0 || h === 0) return;

  // ECMA-376 Part 4 §19.1.2.23 `<v:textpath>` — a WordArt text watermark. It
  // draws stretched, rotated, semi-transparent text filling the shape box
  // INSTEAD of a fill/stroke panel + body text, then returns.
  if (shape.textPath && shape.textPath.string.length > 0) {
    drawWatermarkTextPath(
      ctx as CanvasRenderingContext2D,
      shape.textPath,
      x, y, w, h,
      shape.rotation ?? 0,
      shapeFillColor(shape.fill),
      shape.fillOpacity ?? 1,
      state.fontFamilyClasses,
    );
    return;
  }

  const rot = shape.rotation ?? 0;
  const flipH = shape.flipH ?? false;
  const flipV = shape.flipV ?? false;
  ctx.save();
  // §20.1.7.6 — rotate then mirror about the shape centre, matching the pptx
  // renderer. Applying flip via the canvas transform keeps the body path, the
  // connector arrow-head position, and its direction consistent (a flipped
  // connector swaps which tip carries the head/tail end).
  if (rot !== 0 || flipH || flipV) {
    ctx.translate(x + w / 2, y + h / 2);
    if (rot !== 0) ctx.rotate((rot * Math.PI) / 180);
    if (flipH) ctx.scale(-1, 1);
    if (flipV) ctx.scale(1, -1);
    ctx.translate(-(x + w / 2), -(y + h / 2));
  }
  // Dispatch to the shared spec-driven preset engine when the geometry is a
  // known <a:prstGeom> preset, mirroring the pptx renderer. `arc` (ECMA-376
  // §20.1.10.56 ST_ShapeType "arc") goes through the engine too: its
  // presetShapeDefinitions geometry is two <path>s — path 0 (stroke="false")
  // fills the pie wedge (arc + lnTo centre + close) and path 1 (fill="none")
  // strokes the open arc edge — which the engine honours per-path. The legacy
  // buildShapePath fallback could only draw the open arc, so filling it
  // auto-closed into a chord; arc was excluded here to dodge that, but the
  // engine renders it faithfully now. custGeom (no presetGeometry, subpaths
  // only) still falls through to buildCustomPath.
  const geom = shape.presetGeometry?.toLowerCase() ?? '';
  const usePresetEngine =
    !!shape.presetGeometry && hasPreset(geom);

  const adj = shape.adjValues ?? [];
  const fillStyle = resolveFill(shape.fill, ctx as CanvasRenderingContext2D, x, y, w, h);

  // Build a core Stroke so dash / line-end handling matches the pptx path.
  // `width` is in pt and `scale` is px/pt, so `width * scale` is px — the
  // same convention core's applyStroke / drawArrowHead expect.
  const coreStroke: Stroke | null =
    shape.stroke && (shape.strokeWidth ?? 0) > 0
      ? {
          color: shape.stroke,
          width: shape.strokeWidth ?? 0,
          dashStyle: shape.strokeDash ?? undefined,
          lineCap: shape.strokeCap ?? undefined,
          headEnd: lineEndToArrowEnd(shape.headEnd),
          tailEnd: lineEndToArrowEnd(shape.tailEnd),
        }
      : null;
  const strokeCb = coreStroke
    ? () => {
        applyStroke(ctx as CanvasRenderingContext2D, coreStroke, scale);
        ctx.stroke();
      }
    : null;

  if (usePresetEngine) {
    renderPresetShape(
      ctx as CanvasRenderingContext2D,
      geom, x, y, w, h,
      [
        adj[0] ?? null, adj[1] ?? null, adj[2] ?? null, adj[3] ?? null,
        adj[4] ?? null, adj[5] ?? null, adj[6] ?? null, adj[7] ?? null,
      ],
      fillStyle, strokeCb,
      // docx shapes carry no shadow state, so the clear-shadow hook is a no-op.
      () => {},
      // A retractable connector leader is re-stroked retracted below; suppress
      // the preset engine's full-length leader stroke to avoid a double line /
      // a cap poking through the arrow tip.
      isRetractableLeader && (coreStroke?.headEnd || coreStroke?.tailEnd)
        ? { skipTrailingStroke: true }
        : undefined,
    );
  } else {
    ctx.beginPath();
    if (shape.presetGeometry) {
      buildShapePath(
        ctx as CanvasRenderingContext2D,
        shape.presetGeometry,
        x, y, w, h,
        adj[0] ?? null,
        adj[1] ?? null,
        adj[2] ?? null,
        adj[3] ?? null,
      );
    } else {
      buildCustomPath(ctx as CanvasRenderingContext2D, shape.subpaths, x, y, w, h);
    }
    if (fillStyle) {
      ctx.fillStyle = fillStyle;
      ctx.fill();
    }
    if (strokeCb) strokeCb();
  }

  // Line-end decorations (ECMA-376 §20.1.8.3). Connector/line presets and the
  // callout family both expose head/tail tips with a well-defined tangent; for
  // callouts these decorate the leader line (the geometry's trailing path), not
  // the text rectangle or accent bar. The preset engine does not draw line ends,
  // so this runs whether or not the body went through it. Gate on connector /
  // callout presets only: getConnectorAnchors resolves the last path of any
  // preset, so an arbitrary filled shape carrying an <a:ln> head/tail end would
  // otherwise get spurious arrow heads.
  if (coreStroke && (coreStroke.headEnd || coreStroke.tailEnd) && (isLineGeom || isCalloutGeom)) {
    const anchors = getConnectorAnchors(
      preset, x, y, w, h,
      shape.adjValues ?? [],
    );
    if (anchors) {
      ctx.setLineDash([]);
      // Re-stroke the leader retracted from any filled decoration so the line
      // stops at the arrow base instead of poking through its tip (Word /
      // PowerPoint behaviour). Straight/bent only; curved keeps its preset leader.
      if (isRetractableLeader && anchors.vertices.length >= 2) {
        const pts = anchors.vertices.map((v) => ({ x: v.x, y: v.y }));
        if (coreStroke.tailEnd) {
          const r = lineEndRetract(coreStroke.tailEnd, coreStroke, scale);
          pts[pts.length - 1] = retractLineEndpoint(pts[pts.length - 1], pts[pts.length - 2], r);
        }
        if (coreStroke.headEnd) {
          const r = lineEndRetract(coreStroke.headEnd, coreStroke, scale);
          pts[0] = retractLineEndpoint(pts[0], pts[1], r);
        }
        applyStroke(ctx as CanvasRenderingContext2D, coreStroke, scale);
        ctx.beginPath();
        ctx.moveTo(pts[0].x, pts[0].y);
        for (let i = 1; i < pts.length; i++) ctx.lineTo(pts[i].x, pts[i].y);
        ctx.stroke();
      }
      if (coreStroke.tailEnd) {
        drawArrowHead(ctx as CanvasRenderingContext2D, anchors.end.x, anchors.end.y, anchors.end.angle, coreStroke.tailEnd, coreStroke, scale);
      }
      if (coreStroke.headEnd) {
        drawArrowHead(ctx as CanvasRenderingContext2D, anchors.start.x, anchors.start.y, anchors.start.angle, coreStroke.headEnd, coreStroke, scale);
      }
    }
  }
  ctx.restore();

}

/**
 * Resolve an anchor image's page-space box origin and dist* padding (px), shared
 * by registerAnchorFloats (wrap floats) and renderAnchorImages (wrapNone images).
 *
 * X: margin-relative offsets add section.marginLeft (ECMA-376 §20.4.3.4
 * relativeFrom="margin"); otherwise anchorXPt is already page-absolute.
 * Y: paragraph-relative offsets add `paraBaseY`; otherwise page-absolute. The
 * caller supplies `paraBaseY` = the paragraph's pre-spaceBefore TOP for ALL
 * paragraph-relative floats — wrap and wrapNone alike (ECMA-376 §20.4.3.5: a
 * `positionV relativeFrom="paragraph"` float is positioned relative to the
 * paragraph that contains the anchor, i.e. its top edge before spaceBefore).
 * Page-level floats pass 0 (resolveAnchorY ignores paraBaseY for them). This is
 * the box origin BEFORE any overlap displacement; resolveFloatOverlap runs on
 * top of it for floats.
 *
 * Exported under a `_test` alias for the anchor-image relativeFrom wiring test
 * (the public renderer entry points consume the box internally; pin the
 * positionH/V → xContainer/yContainer plumbing at this seam).
 */
export const __test_resolveAnchorBox = (
  img: ImageRun,
  state: RenderState,
  paraBaseY: number,
): { x: number; y: number; w: number; h: number; dl: number; dr: number; dt: number; db: number } =>
  resolveAnchorBox(img, state, paraBaseY);

/** Exported for the vertical shape-anchor test (ECMA-376 §17.6.20 + §20.4.3.x,
 *  issue #988 ②): pins the physical-page resolution (and logical projection) of
 *  an anchored SHAPE's positionH/V on a vertical (tbRl) page. */
export const __test_resolveShapeBox = (
  shape: ShapeRun,
  state: RenderState,
  paragraphTopPx: number,
): { x: number; y: number; w: number; h: number } =>
  resolveShapeBox(shape, state, paragraphTopPx);

/** Exported for the vertical header/footer test (ECMA-376 §17.6.20 + §17.10.1,
 *  issue #988): pins the inverse-of-`verticalLayoutSection` page/margin mapping a
 *  vertical section's HORIZONTAL header/footer are laid out in. */
export const __test_physicalLayoutSection = (logical: SectionProps): SectionProps =>
  physicalLayoutSection(logical);
export const __test_verticalLayoutSection = (phys: SectionProps): SectionProps =>
  verticalLayoutSection(phys);

/** Exported for the page-anchor pre-scan test (ECMA-376 §20.4.3.2/§20.4.3.5):
 *  drives {@link preRegisterPageFloats} from a unit test against a stub
 *  RenderState so we can pin which paragraphs get pre-registered and that
 *  duplicate calls are idempotent. */
export const __test_preRegisterPageFloats = (
  body: readonly BodyElement[],
  startIdx: number,
  state: RenderState,
): void => preRegisterPageFloats(body, startIdx, state);

/** Exported for the page-anchor pre-scan test — pins the
 *  {paragraph,line,character} ⇒ paragraph-local vs everything-else ⇒ page-level
 *  classification (ECMA-376 §20.4.3.5 ST_RelFromV). */
export const __test_isPageLevelAnchorY = (
  rf: string | null | undefined,
  fromPara: boolean,
): boolean => isPageLevelAnchorY(rf, fromPara);

/** Exported for the table-layout reuse test — resolves a table's px column
 *  widths / row heights through the production {@link computeTableLayout}, so the
 *  test can drive the reuse gate against a stamped element and a stub RenderState. */
export const __test_computeTableLayout = (
  table: DocTable,
  contentWPx: number,
  state: RenderState,
): { colWidths: number[]; tableW: number; rowHeights: number[] } =>
  computeTableLayout(table, contentWPx, state);

/** Exported for the chart-canvas-state-leak regression test (#766): drives the
 *  exact call site (line ~5807) that invokes the shared core `renderChart`
 *  for an inline `<c:chart>` segment, so a unit test can assert that a
 *  `fillText` issued on the SAME ctx right after a chart segment is not left
 *  center-aligned / mis-baselined by chart-internal state that used to leak
 *  past `renderChart` (it now wraps its body in save/restore). */
export const __test_renderInlineImage = (
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  seg: LayoutImageSeg,
  x: number,
  baseline: number,
  scale: number,
  images: Map<string, DecodedImage>,
  vertical = false,
): void => renderInlineImage(ctx, seg, x, baseline, scale, images, vertical);

/** ECMA-376 §17.6.20 + §20.4.3.x — a RenderState view whose page/margin geometry
 *  is the PHYSICAL (un-rotated) page, used to resolve a DrawingML anchor's
 *  `<wp:positionH/V>` against the physical page for a vertical (tbRl) section
 *  (Word places the drawing layer independently of the text-flow rotation). Only
 *  the geometry fields `xContainer`/`yContainer`/`resolveAnchorX`/`resolveAnchorY`
 *  read are overridden (scale, page size, margins, `pageH`); everything else is
 *  the live logical state. Callers map the resolved physical box back into the
 *  logical layout frame with {@link physicalToLogicalAnchorBox}. */
function physicalAnchorState(state: RenderState): RenderState {
  const p = state.verticalPhys;
  if (!p) return state;
  return {
    ...state,
    pageWidth: p.pageWidth,
    marginLeft: p.marginLeft,
    marginRight: p.marginRight,
    marginTop: p.marginTop,
    marginBottom: p.marginBottom,
    // yContainer reads `pageH` (px) for the page-relative bands; the physical
    // page height in px is `pageHeight(pt) * scale`.
    pageH: p.pageHeight * state.scale,
  };
}

/** ECMA-376 §17.6.20 + §20.4.3.x (issue #988 ②/④) — a RenderState view whose
 *  geometry AND text flags are PHYSICAL, for content that stays UPRIGHT inside a
 *  vertical (tbRl) section: anchored shapes and block tables. Word resolves and
 *  paints these against the un-rotated physical page — cell/label text is
 *  horizontal — so on top of {@link physicalAnchorState}'s page/margin un-swap
 *  this view also re-points the content band at the physical margins and clears
 *  the vertical flags (no per-glyph counter-rotation, no +90° text-layer
 *  transform, `resolveShapeBox`/`resolveAnchorBox` take their horizontal path).
 *  `floats` is fresh: the live float set is in LOGICAL flow coordinates and must
 *  not leak into a physical-frame layout (and vice-versa). The front-paint
 *  session is cleared so a nested front float paints in place, inside the
 *  counter-rotated physical frame its geometry was resolved in. */
function verticalPhysicalContentState(state: RenderState): RenderState {
  const p = state.verticalPhys;
  if (!p) return state;
  return {
    ...physicalAnchorState(state),
    contentX: p.marginLeft * state.scale,
    contentW: (p.pageWidth - p.marginLeft - p.marginRight) * state.scale,
    verticalCJK: false,
    verticalAllRotated: false,
    verticalPhys: undefined,
    floats: [],
  };
}

type AnchorBoxSource = Pick<ImageRun,
  | 'widthPt' | 'heightPt'
  | 'anchorXPt' | 'anchorYPt'
  | 'anchorXFromMargin' | 'anchorYFromPara'
  | 'anchorXAlign' | 'anchorYAlign'
  | 'anchorXRelativeFrom' | 'anchorYRelativeFrom'
  | 'distTop' | 'distBottom' | 'distLeft' | 'distRight'
>;

function resolveAnchorBox(
  img: AnchorBoxSource,
  state: RenderState,
  paraBaseY: number,
): { x: number; y: number; w: number; h: number; dl: number; dr: number; dt: number; db: number } {
  const scale = state.scale;
  const w = img.widthPt * scale;
  const h = img.heightPt * scale;
  const dl = (img.distLeft   ?? 0) * scale;
  const dr = (img.distRight  ?? 0) * scale;
  const dt = (img.distTop    ?? 0) * scale;
  const db = (img.distBottom ?? 0) * scale;
  // ECMA-376 §20.4.3.1 wp:align — when positionH/V carry <wp:align>, the
  // renderer aligns the image within its relativeFrom container instead of
  // using the (discarded) posOffset. Mirrors resolveShapeBox (the ShapeRun
  // equivalent): we route X/Y through resolveAnchorX/Y with the image's own
  // box size as the align size. The raw §20.4.3.2/§20.4.3.5
  // `<wp:positionH/V>@relativeFrom` string (e.g. "margin", "topMargin") is
  // threaded through so xContainer/yContainer pick the correct container.
  // Without it a `relativeFrom="margin"` + `align="top"` image would degrade
  // to the page-relative top edge (Y=0 → inside the top margin), which is
  // exactly the sample-11 misplacement before this wire-up. ImageRun carries
  // no pctPos/sizeRel, so those args remain null and the legacy boolean
  // anchorXFromMargin / anchorYFromPara hints still gate page-vs-margin when
  // no raw relativeFrom is present. When align is absent, resolveAnchorX/Y
  // fall back to the offset path.
  if (state.verticalPhys) {
    // ECMA-376 §17.6.20 (tbRl): the anchor's positionH/V are PHYSICAL-page
    // relative (the drawing layer is not rotated with the text flow). Resolve
    // the box in physical space, then project it into the swapped logical layout
    // frame the body text flows in — so the float-exclusion band and the
    // (drawUprightBox-un-swapped) painted image share one geometry. A
    // `paragraph`/`line`-relative positionV anchors from the PHYSICAL TOP of
    // the anchor paragraph's COLUMN (issue #988 batch-3 adjudication ②: Word GT
    // = margin-top + posOffset for a single-column body). That physical y is
    // the column band's logical x start (`state.contentX`; physical y =
    // logical x under the +90° page paint) — NOT the logical flow `paraBaseY`,
    // which lies on the column-progression axis and would rotate the offset.
    const phys = physicalAnchorState(state);
    const px = resolveAnchorX(
      img.anchorXAlign, img.anchorXFromMargin ?? false, img.anchorXPt ?? 0, w, phys,
      img.anchorXRelativeFrom ?? null, null, null,
    );
    const py = resolveAnchorY(
      img.anchorYAlign, img.anchorYFromPara ?? false, img.anchorYPt ?? 0, h, state.contentX, phys,
      img.anchorYRelativeFrom ?? null, null, null,
    );
    const box = physicalToLogicalAnchorBox(px, py, w, h, state.verticalPhys.cssWidthPx);
    // Rotate the dist* padding one quarter-turn with the box: physical top/bottom
    // become logical left/right; physical right/left become logical top/bottom
    // (logical y runs opposite physical x). Symmetric wrapSquare dist is common,
    // but rotate the labels so asymmetric dist stays correct.
    return { x: box.x, y: box.y, w: box.w, h: box.h, dl: dt, dr: db, dt: dr, db: dl };
  }
  const x = resolveAnchorX(
    img.anchorXAlign, img.anchorXFromMargin ?? false, img.anchorXPt ?? 0, w, state,
    img.anchorXRelativeFrom ?? null, null, null,
  );
  const y = resolveAnchorY(
    img.anchorYAlign, img.anchorYFromPara ?? false, img.anchorYPt ?? 0, h, paraBaseY, state,
    img.anchorYRelativeFrom ?? null, null, null,
  );
  return { x, y, w, h, dl, dr, dt, db };
}

/** ECMA-376 §20.4.3.2 / §20.4.3.5 — a `<wp:positionV>` `relativeFrom` value
 *  that resolves the float's Y INDEPENDENTLY of its anchoring paragraph (vs.
 *  `paragraph` / `line` / `character` which resolve against the paragraph's
 *  top). When Y is page-level, Word treats the float as page-positioned: it
 *  is laid out as soon as the page is opened and earlier paragraphs on the
 *  same page wrap around it. {@link preRegisterPageFloats} uses this to
 *  hoist such floats to page-start; paragraph-local Y still flows the legacy
 *  per-paragraph path.
 *
 *  An anchor with NO explicit `<wp:positionV>` (anchorYRelativeFrom absent)
 *  still resolves against the page top via the legacy hint
 *  (`anchorYFromPara=false` ⇒ page-absolute offset), so it qualifies as
 *  page-level too. */
function isPageLevelAnchorY(rf: string | null | undefined, fromPara: boolean): boolean {
  if (rf == null) return !fromPara;
  switch (rf) {
    case 'paragraph':
    case 'line':
    case 'character':
      return false;
    default:
      return true;
  }
}

/** True when this run is a wrap float whose vertical placement is page-level
 *  (independent of source-order paragraph position) — see
 *  {@link isPageLevelAnchorY}. `isWrapFloat` already filters inline images
 *  (their `wrapMode` is undefined) and non-wrapping anchors, so an extra
 *  `anchor` check is redundant here. */
function isPageLevelWrapFloat(run: ImageRun | ChartRun | ShapeRun): boolean {
  if (!isWrapFloat(run.wrapMode)) return false;
  return isPageLevelAnchorY(run.anchorYRelativeFrom ?? null, run.anchorYFromPara ?? false);
}

/** Register floats from a paragraph's anchored images, charts, and shapes.
 *  Images and charts are drawn immediately; anchor shapes are NOT drawn here
 *  (renderAnchorShape paints them separately) — we reserve their float-exclusion
 *  bands so body text wraps around them (ECMA-376 §20.4.2.16/.17).
 *
 *  Page-level floats (positionV relativeFrom ∈ {page, margin, *Margin, column},
 *  ECMA-376 §20.4.3.2/§20.4.3.5) are skipped when this paragraph was already
 *  pre-registered at the current page's start by {@link preRegisterPageFloats}
 *  — re-registering would double-stamp the FloatRect (and re-draw the image).
 *  Paragraph-local floats (`paragraph`/`line`/`character`) keep the per-
 *  paragraph path so their Y stays anchored at this paragraph's top. */
function registerAnchorFloats(para: DocParagraph, state: RenderState, paragraphAnchorY: number): void {
  // One id per registerAnchorFloats call ⇒ one id per paragraph. Floats sharing
  // a paraId (e.g. two side-by-side photos in one paragraph) never displace each
  // other; floats from different paragraphs do (de-facto overlap avoidance).
  const paraId = state.floatParaSeq++;
  const prescanned = state.pageAnchorPrescanned?.has(para) ?? false;
  for (const run of para.runs) {
    if (run.type === 'image') {
      const img = run as unknown as ImageRun;
      if (prescanned && isPageLevelWrapFloat(img)) continue;
      registerImageFloat(img, state, paragraphAnchorY, paraId);
    } else if (run.type === 'chart') {
      const chart = run as unknown as ChartRun;
      if (prescanned && isPageLevelWrapFloat(chart)) continue;
      registerChartFloat(chart, state, paragraphAnchorY, paraId);
    } else if (run.type === 'shape') {
      const shp = run as unknown as ShapeRun;
      if (prescanned && isPageLevelWrapFloat(shp)) continue;
      registerShapeFloat(shp, state, paragraphAnchorY, paraId);
    }
  }
}

/** Pre-scan upcoming body elements at a page-start moment and register any
 *  page-level (positionV relativeFrom ∈ {page, margin, *Margin, column})
 *  wrap floats they carry. Mirrors Word's layout order: page-level floats are
 *  positioned as soon as the page is opened, so paragraphs that PRECEDE the
 *  anchoring paragraph in source order on the same page wrap around them
 *  (ECMA-376 §20.4.3.2/§20.4.3.5 + §20.4.2.16/.17). Each pre-registered
 *  paragraph is recorded in `state.pageAnchorPrescanned` so the main flow's
 *  {@link registerAnchorFloats} skips its page-level runs (avoiding a
 *  duplicate FloatRect / re-drawn image) while still registering its
 *  paragraph-local floats normally.
 *
 *  Bounds: the scan stops at the next forced page boundary that the
 *  paginator/renderer is guaranteed to honor — an explicit `pageBreak`
 *  (§17.18.79 / §17.3.1.20) or a non-continuous `sectionBreak`. Content
 *  overflow may still push paragraphs to later pages mid-scan; the
 *  paginator's `newPage()` resets the float set wholesale, so those
 *  paragraphs get re-pre-scanned on the next page. (Same idempotent flow as
 *  the existing `registerAnchorFloats` post-newPage re-call at the split-
 *  relocation site.) */
function preRegisterPageFloats(
  body: readonly BodyElement[],
  startIdx: number,
  state: RenderState,
): void {
  if (!state.pageAnchorPrescanned) state.pageAnchorPrescanned = new Set();
  for (let j = startIdx; j < body.length; j++) {
    const el = body[j];
    if (!el) continue;
    if (el.type === 'pageBreak') break;
    if (el.type === 'sectionBreak') {
      const sb = el as unknown as { kind?: string };
      if (sb.kind && sb.kind !== 'continuous') break;
      continue;
    }
    if (el.type !== 'paragraph') continue;
    const para = el as unknown as DocParagraph;
    // Skip if already pre-registered (renderer may call this once per page;
    // paginator may re-call after newPage(), but newPage clears the set).
    if (state.pageAnchorPrescanned.has(para)) continue;
    let hasPageLevel = false;
    for (const run of para.runs) {
      if (run.type === 'image') {
        if (isPageLevelWrapFloat(run as unknown as ImageRun)) { hasPageLevel = true; break; }
      } else if (run.type === 'chart') {
        if (isPageLevelWrapFloat(run as unknown as ChartRun)) { hasPageLevel = true; break; }
      } else if (run.type === 'shape') {
        if (isPageLevelWrapFloat(run as unknown as ShapeRun)) { hasPageLevel = true; break; }
      }
    }
    if (!hasPageLevel) continue;
    // Register only the page-level floats from this paragraph. paraY=0 is safe
    // because resolveAnchorY ignores it for page-level relativeFrom containers
    // (anchor-geometry.ts §20.4.3.x). Allocate a fresh paraId so overlap
    // avoidance treats these like any other anchor-paragraph float.
    const paraId = state.floatParaSeq++;
    for (const run of para.runs) {
      if (run.type === 'image') {
        const img = run as unknown as ImageRun;
        if (!isPageLevelWrapFloat(img)) continue;
        registerImageFloat(img, state, 0, paraId);
      } else if (run.type === 'chart') {
        const chart = run as unknown as ChartRun;
        if (!isPageLevelWrapFloat(chart)) continue;
        registerChartFloat(chart, state, 0, paraId);
      } else if (run.type === 'shape') {
        const shp = run as unknown as ShapeRun;
        if (!isPageLevelWrapFloat(shp)) continue;
        registerShapeFloat(shp, state, 0, paraId);
      }
    }
    state.pageAnchorPrescanned.add(para);
  }
}

/** Reserve the float-exclusion rect for one anchored wrap-image and draw the
 *  bitmap immediately (the image is the float). */
function registerImageFloat(
  img: ImageRun,
  state: RenderState,
  paragraphAnchorY: number,
  paraId: number,
): void {
  if (!img.anchor) return;
  if (!isWrapFloat(img.wrapMode)) return;

  const mode: 'square' | 'topAndBottom' =
    img.wrapMode === 'topAndBottom' ? 'topAndBottom' : 'square';

  // Paragraph-relative wrap floats anchor at the pre-spaceBefore paragraph top
  // (paragraphAnchorY), per ECMA-376 §20.4.3.5 — identical to wrapNone images.
  const box = resolveAnchorBox(img, state, paragraphAnchorY);
  const { w, h, dl, dr, dt, db } = box;

  // Overlap avoidance. Spec-mandated part: allowOverlap="false" (ECMA-376
  // §20.4.2.3) REQUIRES repositioning to prevent overlap; "true"/omitted only
  // permits overlap. Default true per §20.4.2.3.
  // Implementation-defined (HEURISTIC, Word-mimicking, no ECMA-376 basis):
  // displacing the later document-order float, the "other paragraphs only"
  // gate under allowOverlap=true, and the right-then-down re-seat using dist
  // padding as the float-to-float gap. See resolveFloatOverlap header.
  const allowOverlap = img.allowOverlap ?? true;
  const key = imageKey(img.imagePath, img.colorReplaceFrom, img.duotone);
  const rect = pushFloatRect(state, {
    x: box.x,
    y: box.y,
    w, h, dl, dr, dt, db,
    kind: 'shape', // DrawingML anchor (§20.4.2.3); not a floating table.
    mode,
    side: img.wrapSide ?? 'bothSides',
    imageKey: key,
    drawn: false,
    paraId,
    avoidOverlap: true,
    allowOverlap,
  });

  if (!state.dryRun) {
    const bmp = state.images.get(key);
    if (bmp) {
      // §20.1.8.6 alphaModFix — multiply the picture's opacity.
      const hasAlpha = img.alpha != null && img.alpha < 1;
      if (hasAlpha) {
        state.ctx.save();
        state.ctx.globalAlpha *= img.alpha as number;
      }
      if (state.verticalCJK) {
        // §17.6.20 (tbRl) — keep the floated image UPRIGHT inside the rotated page.
        drawUprightBox(state.ctx, rect.imageX, rect.imageY, rect.imageW, rect.imageH, (dx, dy, dw, dh) =>
          drawImageRunBitmap(state.ctx, bmp, img, dx, dy, dw, dh),
        );
      } else {
        drawImageRunBitmap(state.ctx, bmp, img, rect.imageX, rect.imageY, rect.imageW, rect.imageH);
      }
      if (hasAlpha) state.ctx.restore();
    }
    rect.drawn = true;
  }
}

/** Reserve the float-exclusion rect for one anchored wrap-chart and paint the
 *  chart at the overlap-resolved box (ECMA-376 §20.4.2.3/.16/.17). */
function registerChartFloat(
  chart: ChartRun,
  state: RenderState,
  paragraphAnchorY: number,
  paraId: number,
): void {
  if (!chart.anchor || !isWrapFloat(chart.wrapMode)) return;

  const box = resolveAnchorBox(chart, state, paragraphAnchorY);
  const { w, h, dl, dr, dt, db } = box;
  if (w <= 0 || h <= 0) return;

  const rect = pushFloatRect(state, {
    x: box.x,
    y: box.y,
    w, h, dl, dr, dt, db,
    kind: 'shape',
    mode: chart.wrapMode === 'topAndBottom' ? 'topAndBottom' : 'square',
    side: chart.wrapSide ?? 'bothSides',
    allowOverlap: chart.allowOverlap ?? true,
    avoidOverlap: true,
    paraId,
    imageKey: '',
    drawn: false,
  });

  if (!state.dryRun) {
    const paint = (x: number, y: number, width: number, height: number): void =>
      renderChart(
        state.ctx as CanvasRenderingContext2D,
        chart.chart,
        { x, y, w: width, h: height },
        state.scale,
      );
    if (state.verticalCJK) {
      // ECMA-376 §17.6.20 (tbRl) — a chart is a graphic, so keep it upright.
      drawUprightBox(state.ctx, rect.imageX, rect.imageY, rect.imageW, rect.imageH, paint);
    } else {
      paint(rect.imageX, rect.imageY, rect.imageW, rect.imageH);
    }
    rect.drawn = true;
  }
}

/** Reserve the float-exclusion rect for one anchored wrap-shape (wps:txbx /
 *  DrawingML wp:anchor shape). The shape is drawn separately by
 *  renderAnchorShape, so here we only push the FloatRect (drawn=true ⇒ the
 *  deferred-image-draw path never tries to paint it). The box is resolved by the
 *  SAME resolveShapeBox the renderer draws with, so the band matches the shape. */
function registerShapeFloat(
  shape: ShapeRun,
  state: RenderState,
  paragraphAnchorY: number,
  paraId: number,
): void {
  if (!isWrapFloat(shape.wrapMode)) return;

  // Match resolveShapeBox's paragraphTopPx convention. resolveAnchorY reads
  // paragraphTopPx only for relativeFrom="paragraph"/"line" (anchorYFromPara);
  // wrap floats anchor at the pre-spaceBefore paragraph top (§20.4.3.5),
  // identical to the image path (resolveAnchorBox uses paragraphAnchorY there).
  const { x, y, w, h } = resolveShapeBox(shape, state, paragraphAnchorY);
  // A degenerate (zero/negative-area) box — e.g. a wrap-flagged line preset —
  // reserves no band; bail like renderAnchorShape skips drawing it.
  if (w <= 0 || h <= 0) return;

  const mode: 'square' | 'topAndBottom' =
    shape.wrapMode === 'topAndBottom' ? 'topAndBottom' : 'square';

  const scale = state.scale;
  const pdl = (shape.distLeft   ?? 0) * scale;
  const pdr = (shape.distRight  ?? 0) * scale;
  const pdt = (shape.distTop    ?? 0) * scale;
  const pdb = (shape.distBottom ?? 0) * scale;
  // §17.6.20 — on a vertical page the box above is the LOGICAL projection of the
  // physically-resolved shape (resolveShapeBox), so rotate the dist* labels one
  // quarter-turn with it, exactly like the image path (resolveAnchorBox):
  // physical top/bottom ↦ logical left/right, physical right/left ↦ logical
  // top/bottom (logical y runs opposite physical x).
  const vertical = !!state.verticalPhys;
  const dl = vertical ? pdt : pdl;
  const dr = vertical ? pdb : pdr;
  const dt = vertical ? pdr : pdt;
  const db = vertical ? pdl : pdb;

  // Overlap avoidance, kept consistent with the image path. Shapes carry no
  // parsed allowOverlap field; the spec default is true (§20.4.2.3), so
  // same-paragraph floats never displace each other and a lone shape is a no-op
  // here — but running it keeps multi-float behavior identical to images.
  pushFloatRect(state, {
    x, y, w, h, dl, dr, dt, db,
    kind: 'shape', // DrawingML wp:anchor shape (§20.4.2.3); not a floating table.
    mode,
    side: shape.wrapSide ?? 'bothSides',
    imageKey: '',
    // The shape is painted by renderAnchorShape, not by the deferred image-draw
    // path; mark it drawn so that path skips it (it has no bitmap to draw).
    drawn: true,
    paraId,
    avoidOverlap: true,
    allowOverlap: true,
  });
}

// ===== Table rendering =====

/** Per-column widths (px), total table width (px), and per-row heights (px,
 *  with the §17.4.85 vMerge-span extension applied) for a table laid out in a
 *  content band `contentWPx` wide. Shared by the block ({@link renderTable}) and
 *  floating ({@link renderFloatTable}) paths so both size the table identically. */
function computeTableLayout(
  table: DocTable,
  contentWPx: number,
  state: RenderState,
  sourceIndex?: number,
): {
  colWidths: number[];
  tableW: number;
  rowContentHeights: number[];
  rowHeights: number[];
} {
  const { scale } = state;
  const contentWPt1 = contentWPx / scale;
  // Body occurrences pass their source index explicitly. Header/footer tables
  // have no body occurrence owner and retain the B1 story calculation below.
  if (state.retainedTableAcquisition && sourceIndex !== undefined) {
    const retained = computeTablePtLayout(state, table, contentWPt1, sourceIndex);
    const colWidths = retained.colWidthsPt.map((width) => width * scale);
    const rowHeights = retained.rowHeightsPt.map((height) => height * scale);
    return {
      colWidths,
      tableW: colWidths.reduce((sum, width) => sum + width, 0),
      rowContentHeights: retained.rowContentHeightsPt.map((height) => height * scale),
      rowHeights,
    };
  }

  // Resolve column widths in pt (autofit by preferred widths, or fixed grid),
  // already scaled to fit the available content width, then convert to px.
  const colWidths = resolveColumnWidths(table, contentWPt1, state).map((w) => w * scale);
  const tableW = colWidths.reduce((s, w) => s + w, 0);

  // Shared ST_HeightRule + §17.4.85 vMerge-span skeleton (resolveTableRowHeights),
  // with the paint pass's px cell measurer. The restart-span extension is part of
  // the skeleton now — calculateRowHeight already excludes restart cells per-row,
  // and the resolver re-measures them via the same callback to grow the last row.
  const rowContentHeights = resolveTableRowContentHeights(table, colWidths, scale, (cell, cellW) =>
    measureCellContentHeightPx(cell, table, cellW, scale, state),
  );
  const rowHeights = applyTableRowBoundaryFootprints(table, rowContentHeights, scale);

  return { colWidths, tableW, rowContentHeights, rowHeights };
}

/** Content height of a table cell laid out at total width `cellW`, in the target
 *  units the caller works in: px when `scale` is the device scale (paint pass),
 *  pt when `scale === 1` (paginator). Cell top/bottom margins plus each content
 *  element measured at `measureCellElementHeight`. Adjacent paragraphs inside the
 *  cell collapse spacing the same way `renderCellContent` does (ECMA-376
 *  §17.3.1.33 contextualSpacing + spaceAfter/spaceBefore overlap = max not sum),
 *  so the measured height matches the painted height.
 *
 *  B2 table stage 1a — this is the SINGLE cell-content measurer for the whole
 *  package. The paginator ({@link computeTableRowHeights}, scale 1), the paint
 *  layout ({@link computeTableLayout}, device scale), and the exported
 *  {@link calculateRowHeight} all resolve a cell's height through here, so a
 *  table's rows can never be sized by two different measurers. Unit-agnostic: it
 *  is the same formula at any `scale`, and at scale 1 it returns exactly the pt
 *  height the device-scale paint pass will produce ÷ scale. */
function measureCellContentHeightPx(
  cell: DocTableCell,
  table: DocTable,
  cellW: number,
  scale: number,
  state: RenderState,
): number {
  const cellState = withTableCellStory(state);
  const cm = effCellMargins(cell, table);
  const contentW = cellW - (cm.left + cm.right) * scale;
  // ECMA-376 §17.4.7 requires every <w:tc> to end with a <w:p>. When the cell's
  // visible content is a nested table, Word emits a trailing empty <w:p/> purely
  // as that syntactic anchor; it carries no ink and does NOT grow the row (Word's
  // outer cell hugs the inner table — sample-11's "table inside a table" outer
  // row measures the inner table height, not inner + the structural mark's line
  // box + its inherited space-before). Drop it from the row-height measurement,
  // exactly as the vAlign block height already does (trimTrailingStructuralMarker).
  // The mark itself is still painted by renderCellContent; being empty it adds no
  // visible content, so excluding it from sizing cannot hide anything.
  const measured = trimTrailingStructuralMarker(cell.content);
  // measureCellElementHeight always includes paragraph spaceBefore plus
  // max(spaceAfter, bottom-border extent) — the same trailing advance the paint
  // pass emits (§17.3.1.7); sumCellContentHeight folds in contextualSpacingAdjust
  // (§17.3.1.9) and the prevSpaceAfter/spaceBefore overlap collapse to match the
  // paint pass's renderCellContent. Spacing is converted from pt to px with `scale`.
  return (cm.top + cm.bottom) * scale + sumCellContentHeight(
    measured,
    (ce) => measureCellElementHeight(cellState, ce, contentW, scale),
    scale,
  );
}

/** Draw all rows of a table whose grid origin is `tableX` (px) and whose top is
 *  `startY` (px), returning the Y just past the last row. Shared by the block
 *  and floating paths. Honors bidiVisual, vMerge span heights, and exact-row
 *  clipping exactly as the original inline loop did. In dryRun, measures cell
 *  content instead of drawing. */
interface TableCellPaintJob {
  cell: DocTableCell;
  x: number;
  y: number;
  w: number;
  h: number;
  edges: CellEdgeFlags;
  clipExact: boolean;
  /** ECMA-376 §17.4.66 — this cell's grid footprint, so the border pass can find
   *  the cells that share each interior gridline (its right/bottom neighbours).
   *  `ci`/`ri` are the logical top-left grid slot; `span` the column span; `lastRi`
   *  the last row the cell occupies (vMerge-span aware). */
  ci: number;
  ri: number;
  span: number;
  lastRi: number;
}

function drawTableRows(
  table: DocTable,
  colWidths: number[],
  tableW: number,
  rowHeights: number[],
  tableX: number,
  startY: number,
  state: RenderState,
): number {
  const { scale, dryRun } = state;
  // ECMA-376 §17.4.1 `<w:bidiVisual>`: lay the grid columns right-to-left, so
  // logical column 0 sits at the table's RIGHT edge and indices advance
  // leftward. We mirror by POSITION arithmetic (not canvas transform): a cell
  // spanning [ci, ci+span) gets physical left x = tableX + tableW − (offset of
  // its right grid edge). Cell borders are mirrored too — a cell's logical
  // left/right border specs swap physical sides (its "start" edge is on the
  // right). gridSpan still consumes the same logical columns; only the mapping
  // from logical column offset to a physical x flips.
  const mirror = table.bidiVisual === true;

  // ECMA-376 §17.4.66 (border-collapse): a shared gridline must sit ON TOP of
  // every cell fill. Painting cell-by-cell (fill→border, per cell) let the next
  // column's background fill cover the half of the vertical border the previous
  // column had just drawn; with alternating row banding (e.g. Medium List 2)
  // this made a shared vertical rule look like its thickness changed row to row.
  // So walk the grid ONCE to collect every cell's paint box, then paint in TWO
  // passes: all backgrounds + content first, all borders second.
  const jobs: TableCellPaintJob[] = [];
  // ECMA-376 §17.4.66 — grid-slot → job index occupancy, so an interior edge can
  // look up the adjacent cell. A vMerge/colSpan cell fills every slot it covers
  // (its restart job index), so a neighbour query on any slot resolves to the
  // owning job. Continue (vMerge=false) cells are covered by their restart job.
  const occupancy: number[][] = table.rows.map(() => new Array<number>(colWidths.length).fill(-1));

  let y = startY;
  for (let ri = 0; ri < table.rows.length; ri++) {
    const row = table.rows[ri];
    const rowH = rowHeights[ri];
    let ci = rowGridBefore(row, colWidths.length);
    let x = tableX + colWidths.slice(0, ci).reduce((sum, width) => sum + width, 0);

    for (const cell of row.cells) {
      const span = Math.min(cell.colSpan, colWidths.length - ci);
      const cellW = colWidths.slice(ci, ci + span).reduce((s, w) => s + w, 0);
      // Physical left edge of this cell. LTR: cumulative from the left (`x`).
      // bidiVisual: place so logical column 0 is rightmost — the cell's left
      // edge is the table's right edge minus the offset of its trailing grid
      // line (sum of widths up to and including this span).
      const leadX = mirror ? tableX + tableW - (x - tableX) - cellW : x;

      if (cell.vMerge === false) {
        // continue cell — content is rendered by its restart partner.
      } else {
        // ECMA-376 §17.4.85: a vMerge=restart cell visually occupies the full
        // merged span; use the sum of row heights for its render box.
        let drawH = rowH;
        let lastRowOfCell = ri;
        if (cell.vMerge === true) {
          const endRi = findMergeEndRow(table, ri, ci);
          lastRowOfCell = endRi;
          drawH = 0;
          for (let rj = ri; rj <= endRi; rj++) drawH += rowHeights[rj];
        }
        // ECMA-376 §17.4.38/§17.4.39: classify which physical edges of this cell
        // are the table's OUTER edges (vs. interior gridlines) from its grid
        // position so resolveCellEdges can pick table.top/bottom/left/right vs.
        // table.insideH/insideV. `leftCol`/`rightCol` are the LOGICAL columns
        // (gridSpan-aware); the renderer flips them for bidiVisual via `mirror`.
        const edges: CellEdgeFlags = {
          topRow: ri === 0,
          bottomRow: lastRowOfCell === table.rows.length - 1,
          leftCol: ci === 0,
          rightCol: ci + span === colWidths.length,
        };
        // ECMA-376 §17.4.81: an exact row height is honored verbatim and
        // content taller than the row is clipped to the row box (Word clips;
        // we would otherwise overflow into neighboring rows). A vMerge=restart
        // cell spans multiple rows, so it is never governed by a single row's
        // exact height — only single-row cells clip.
        const clipExact = row.rowHeightRule === 'exact' && cell.vMerge !== true;
        if (dryRun) measureCellContent(cell, table, cellW, scale, state);
        else {
          const jobIndex = jobs.length;
          jobs.push({ cell, x: leadX, y, w: cellW, h: drawH, edges, clipExact, ci, ri, span, lastRi: lastRowOfCell });
          // ECMA-376 §17.4.66 — record this cell's grid footprint so interior-edge
          // neighbour lookups resolve to it. A vMerge=restart cell owns every row it
          // spans; a colSpan cell owns every column in its span.
          for (let rj = ri; rj <= lastRowOfCell && rj < occupancy.length; rj++) {
            for (let cj = ci; cj < ci + span && cj < colWidths.length; cj++) {
              occupancy[rj][cj] = jobIndex;
            }
          }
        }
      }

      x += cellW;
      ci += span;
    }

    y += rowH;
  }

  // Pass 1: backgrounds + content. Pass 2: borders, so a border is never
  // overpainted by a neighbouring cell's fill. `mirror` only swaps which
  // physical side a logical border maps to, so it is consulted in the border
  // pass alone.
  for (const j of jobs) {
    renderCell(j.cell, table, j.x, j.y, j.w, j.h, state, j.clipExact);
  }

  // ECMA-376 §17.4.66 — adjacent-cell border conflict resolution. Each SHARED
  // interior gridline is drawn ONCE with the §17.4.66 winner (weight → precedence
  // → luminance → reading order), instead of both neighbours painting it and the
  // later one winning. Ownership convention so every line is drawn exactly once:
  //   • outer table edges → the single bordering cell draws its own resolved spec;
  //   • interior VERTICAL line → the LEFT cell owns it (drawn as its right edge),
  //     resolved against the RIGHT neighbour's left edge;
  //   • interior HORIZONTAL line → the ABOVE cell owns it (drawn as its bottom
  //     edge), resolved against the BELOW neighbour's top edge.
  // A cell's own top/left INTERIOR edges are therefore not drawn by the cell — the
  // neighbour that owns that gridline consults this cell's spec as the opponent.
  const ctx = state.ctx;
  const dpr = state.dpr;
  // ECMA-376 §17.4.66 (#815) — physical positions of the grid-line boundaries so a
  // shared interior edge can be SUBDIVIDED at neighbour-cell boundaries: a merged
  // cell (gridSpan/vMerge) faces several finer neighbours along one edge, and each
  // sub-segment must be resolved against its OWN opposing cell. colOff/rowOff are
  // LTR cumulative sizes; colBoundaryX folds the bidiVisual flip (§17.4.1), while
  // row positions are never mirrored.
  const colOff: number[] = [0];
  for (const cw of colWidths) colOff.push(colOff[colOff.length - 1] + cw);
  const rowOff: number[] = [0];
  for (const rh of rowHeights) rowOff.push(rowOff[rowOff.length - 1] + rh);
  const colBoundaryX = (c: number): number =>
    mirror ? tableX + tableW - colOff[c] : tableX + colOff[c];
  const rowBoundaryY = (r: number): number => startY + rowOff[r];
  for (const j of jobs) {
    const { x, y, w, h } = j;
    const own = resolveCellEdges(j.cell.borders, table.borders, j.edges, mirror);
    // ECMA-376 §17.4.66 / §17.4.85: a vertical merge is serialized as a
    // restart cell followed by continuation cells. Its bottom boundary belongs
    // to the LAST continuation cell, whose tcBorders can explicitly differ from
    // the restart cell (commonly `bottom="nil"` to leave an adjacent area open).
    // Using only the restart borders incorrectly falls through to table insideH
    // and paints a rule that Word suppresses.
    const terminalCell = j.lastRi > j.ri
      ? (cellAtGridColumn(table.rows[j.lastRi], j.ci, colWidths.length) ?? j.cell)
      : j.cell;
    const ownBottom = terminalCell === j.cell
      ? own.bottom
      : resolveCellEdges(
          terminalCell.borders,
          table.borders,
          { ...j.edges, topRow: false },
          mirror,
        ).bottom;
    // `own.left`/`own.right` are already PHYSICAL (resolveCellEdges folded the
    // bidiVisual swap into the spec). The OUTER-vs-interior GATE must be physical
    // too: under mirror a cell's physical-left edge is the logical RIGHT edge, so
    // the physical outer-left flag is `edges.rightCol` (and vice versa). The
    // physical-right NEIGHBOUR sits at the grid slot on that physical side —
    // logical `ci + span` in LTR, logical `ci - 1` under mirror.
    const physLeftOuter = mirror ? j.edges.rightCol : j.edges.leftCol;
    const physRightOuter = mirror ? j.edges.leftCol : j.edges.rightCol;
    const physLeftCi = mirror ? j.ci + j.span : j.ci - 1;
    const physRightCi = mirror ? j.ci - 1 : j.ci + j.span;

    // TOP: the outer top row draws its own top. Interior tops are normally owned
    // by the cell above (its bottom), but a row omission can leave individual
    // grid slots empty. No neighbour owns those sub-segments, so this cell must
    // paint its own resolved top there.
    if (j.edges.topRow) {
      const spec = paintable(own.top?.spec ?? null);
      if (spec) drawBorderLine(ctx, x, y, x + w, y, spec, scale, dpr);
    } else {
      const spec = paintable(own.top?.spec ?? null);
      if (spec) {
        let cj = j.ci;
        while (cj < j.ci + j.span) {
          const idx = occupancy[j.ri - 1]?.[cj] ?? -1;
          let cEnd = cj + 1;
          while (
            cEnd < j.ci + j.span &&
            (occupancy[j.ri - 1]?.[cEnd] ?? -1) === idx
          ) cEnd++;
          if (idx < 0) {
            drawBorderLine(
              ctx,
              colBoundaryX(cj),
              y,
              colBoundaryX(cEnd),
              y,
              spec,
              scale,
              dpr,
            );
          }
          cj = cEnd;
        }
      }
    }
    // PHYSICAL LEFT: outer-left draws its own edge. An interior left is normally
    // owned by the physically-left neighbour, except where gridBefore/gridAfter
    // leaves that slot empty for all or part of a vertically merged cell.
    if (physLeftOuter) {
      const spec = paintable(own.left?.spec ?? null);
      if (spec) drawBorderLine(ctx, x, y, x, y + h, spec, scale, dpr);
    } else {
      const spec = paintable(own.left?.spec ?? null);
      if (spec) {
        let rj = j.ri;
        while (rj <= j.lastRi) {
          const idx = occupancy[rj]?.[physLeftCi] ?? -1;
          let rEnd = rj;
          while (
            rEnd + 1 <= j.lastRi &&
            (occupancy[rEnd + 1]?.[physLeftCi] ?? -1) === idx
          ) rEnd++;
          if (idx < 0) {
            drawBorderLine(
              ctx,
              x,
              rowBoundaryY(rj),
              x,
              rowBoundaryY(rEnd + 1),
              spec,
              scale,
              dpr,
            );
          }
          rj = rEnd + 1;
        }
      }
    }
    // BOTTOM: outer bottom → own spec; interior → resolve vs each below neighbour's
    // top and draw the winner (this ABOVE cell owns the shared horizontal line).
    if (j.edges.bottomRow) {
      // Mid-row page cut (fidelity round, measured ground truth): the cut is
      // a SHARED horizontal edge between this piece's cell bottom and the
      // continuation piece's cell top on the next page — resolve it with the
      // ordinary §17.4.66 conflict against a SYNTHETIC continuation sibling
      // built from the SAME source-row cell specs, whose top resolves as the
      // next slice-table's OUTER top (cell.top ?? table.top). This explains
      // both measured classes: none ∨ single → single (the form label column
      // with no bottom border still shows the full-width cut rule), and a
      // borderless table draws nothing. The sibling exists only here — it is
      // never inserted into pagination, fragments, or occupancy — and the
      // continuation piece still draws its own outer top on its page.
      // Row-boundary cuts carry no marker and keep the plain outer bottom.
      let spec: BorderSpec | null;
      const cutRow = table.rows[j.lastRi] as DocTableRow & { pageCutBottom?: boolean };
      if (cutRow?.pageCutBottom === true) {
        const siblingTop = resolveCellEdges(
          j.cell.borders,
          table.borders,
          { ...j.edges, topRow: true },
          mirror,
        ).top;
        spec = resolveSharedEdge(ownBottom, siblingTop);
      } else {
        spec = paintable(ownBottom?.spec ?? null);
      }
      if (spec) drawBorderLine(ctx, x, y + h, x + w, y + h, spec, scale, dpr);
    } else if ((table.rows[j.lastRi] as DocTableRow & { pageCutBottom?: boolean })?.pageCutBottom === true) {
      // Intra-row page cut whose CONTINUATION piece shares this page. When a tall
      // row is split, the paginator can pack a leading piece and its continuation
      // onto one page (measured private fixture sample-33 p.3: two consecutive
      // tall rows each split, a continuation piece of each landing on one page).
      // The leading piece's bottom is then an INTERIOR horizontal edge (not the
      // table's outer bottom, which the `j.edges.bottomRow` branch above handles).
      // The pieces are one continuous flow, so Word leaves the cut OPEN — draw
      // NOTHING here, rather than resolving §17.4.66 against the piece below and
      // painting the Table-Grid insideH. (The true page-end cut — the LAST piece
      // on the page — keeps its rule via the outer-bottom branch, whose
      // synthetic-sibling resolution is unchanged.)
    } else {
      // ECMA-376 §17.4.66 (#815) — the shared horizontal edge below this cell may
      // face SEVERAL finer below-cells (this cell is wider via gridSpan). Subdivide
      // the edge at the below-cells' column boundaries and resolve EACH sub-segment
      // against its OWN below neighbour, drawing a per-segment winner rather than
      // resolving the whole edge against the span-origin neighbour alone.
      const belowRi = j.lastRi + 1;
      let cj = j.ci;
      while (cj < j.ci + j.span) {
        const idx = occupancy[belowRi][cj];
        let cEnd = cj + 1;
        while (cEnd < j.ci + j.span && occupancy[belowRi][cEnd] === idx) cEnd++;
        const below = neighbourJob(jobs, occupancy, belowRi, cj);
        const belowEdges = below
          ? resolveCellEdges(below.cell.borders, table.borders, below.edges, mirror)
          : null;
        const spec = resolveSharedEdge(ownBottom, belowEdges?.top ?? null);
        if (spec) drawBorderLine(ctx, colBoundaryX(cj), y + h, colBoundaryX(cEnd), y + h, spec, scale, dpr);
        cj = cEnd;
      }
    }
    // PHYSICAL RIGHT: outer-right → own spec; interior → resolve vs the physically-
    // right neighbour's left and draw the winner (this cell owns the shared vertical
    // line as its physical right edge — so each line is drawn once).
    if (physRightOuter) {
      const spec = paintable(own.right?.spec ?? null);
      if (spec) drawBorderLine(ctx, x + w, y, x + w, y + h, spec, scale, dpr);
    } else {
      // ECMA-376 §17.4.66 (#815) — a vMerge cell's physical-right edge may face
      // SEVERAL finer right-neighbours down the rows it spans. Subdivide the edge at
      // those neighbours' row boundaries and resolve EACH sub-segment against its OWN
      // neighbour's facing (left) edge, drawing a per-segment winner.
      let rj = j.ri;
      while (rj <= j.lastRi) {
        const idx = occupancy[rj][physRightCi];
        let rEnd = rj;
        while (rEnd + 1 <= j.lastRi && occupancy[rEnd + 1][physRightCi] === idx) rEnd++;
        const right = neighbourEdges(jobs, occupancy, rj, physRightCi, mirror, table.borders);
        const spec = resolveSharedEdge(own.right, right?.left ?? null);
        if (spec) drawBorderLine(ctx, x + w, rowBoundaryY(rj), x + w, rowBoundaryY(rEnd + 1), spec, scale, dpr);
        rj = rEnd + 1;
      }
    }
  }
  return y;
}

/** Resolve the neighbour cell at grid slot (ri, ci) and return its resolved
 *  edges, or `null` when the slot is empty/out of range. */
function neighbourEdges(
  jobs: ReadonlyArray<TableCellPaintJob>,
  occupancy: number[][],
  ri: number,
  ci: number,
  mirror: boolean,
  table: TableBorders,
): ResolvedCellEdges | null {
  const nb = neighbourJob(jobs, occupancy, ri, ci);
  return nb ? resolveCellEdges(nb.cell.borders, table, nb.edges, mirror) : null;
}

function neighbourJob(
  jobs: ReadonlyArray<TableCellPaintJob>,
  occupancy: number[][],
  ri: number,
  ci: number,
): TableCellPaintJob | null {
  if (ri < 0 || ri >= occupancy.length) return null;
  if (ci < 0 || ci >= occupancy[ri].length) return null;
  const idx = occupancy[ri][ci];
  if (idx < 0) return null;
  return jobs[idx] ?? null;
}

/** ECMA-376 §17.4.66 — pick the winning border for a shared interior edge from
 *  the two neighbouring cells' resolved edges, then reduce it to a paintable
 *  {@link BorderSpec} (nil/none ⇒ null). `a` is the owning (reading-order-first)
 *  side. */
function resolveSharedEdge(
  a: { spec: BorderSpec; source: 'cell' | 'table' } | null,
  b: { spec: BorderSpec; source: 'cell' | 'table' } | null,
): BorderSpec | null {
  const winner = resolveBorderConflict(a, b);
  return winner ? paintable(winner.spec) : null;
}

/**
 * Render a FLOATING table (ECMA-376 §17.4.57 `<w:tblpPr>`). Like a `<w:framePr>`
 * frame, it is OUT OF FLOW: drawn at an absolute (anchor-relative) position and
 * consuming ZERO flow height, so the following content begins where the table's
 * anchor paragraph sat. A wrap-exclusion FloatRect is registered (§17.4.57
 * *FromText padding, §17.4.56 overlap) so the body text flows around it.
 *
 * Mirrors {@link renderFrameParagraph}: save contentX/contentW/y, redirect the
 * flow geometry to the resolved box, draw the rows, then RESTORE the in-flow
 * state.y (the float adds no flow height). In dryRun the rows are not drawn but
 * the float is still registered so wrap estimates see the band.
 */
function renderFloatTable(table: DocTable, state: RenderState): void {
  const tp = effectiveTablePositioning(table);
  if (!tp) throw new Error('Floating table paint requires effective positioning');
  const inFlowY = state.y;
  const savedX = state.contentX;
  const savedW = state.contentW;

  // Lay the table out in its anchor column's content band, then place its box.
  // tableW is the ACTUAL rendered width (sum of column widths), so the FloatRect
  // exclusion matches the painted table exactly (#513 column integrity: for
  // horzAnchor="text" the box.x derives from the column band via
  // frameXContainer, so the wrap stays inside this column).
  const { colWidths, tableW, rowHeights } = computeTableLayout(table, state.contentW, state);
  const tableH = rowHeights.reduce((s, h) => s + h, 0);
  const box = computeFloatTableBox(tp, state, inFlowY, tableW, tableH);
  const side = floatTableWrapSide(box, state);

  // Redirect the flow geometry to the float box and draw the rows there. The
  // table is positioned absolutely; its grid origin is box.x (no jc — a floating
  // table's position is dictated entirely by tblpPr, §17.4.57).
  state.contentX = box.x;
  state.contentW = tableW;
  drawTableRows(table, colWidths, tableW, rowHeights, box.x, box.y, state);
  state.contentX = savedX;
  state.contentW = savedW;

  // Restore the in-flow cursor: a floating table consumes NO body flow height
  // (§17.4.57 — out of flow), so the following content spaces as if it weren't
  // here. (renderFrameParagraph does the same for a frame.)
  state.y = inFlowY;

  registerTableFloat(box, tp, state, side, table.overlap !== 'never');
}

function renderTable(table: DocTable, state: RenderState): void {
  // ECMA-376 §17.4.57 + [MS-OI29500] §2.1.162: divert only an
  // effectively positioned table to the out-of-flow path; Word can ignore an
  // authored tblpPr, so lexical presence alone is not this branch condition.
  if (!tableParticipatesInOrdinaryFlow(table)) {
    renderFloatTable(table, state);
    return;
  }

  // ECMA-376 §17.6.20 + §17.4.80 (issue #988 batch-3 adjudication ④): a block
  // table inside a vertical (tbRl) section renders UPRIGHT — its cells do NOT
  // inherit the section text direction (cell text is horizontal), a fixed
  // `tcW` is a PHYSICAL width, and `trHeight` exact/auto clip/grow along the
  // physical vertical axis. Lay the table out with the PHYSICAL state view and
  // paint it inside the inverse of the +90° page transform (the same physical
  // frame the header/footer and anchored-shape paths re-enter). Its placement
  // in the vertical flow: the physical TOP edge sits at the column axis start
  // (the logical x band start, contentX ⇒ the physical top content margin —
  // matching Word GT, both fixture tables pinned at the top margin) and the
  // block advances the flow by its PHYSICAL WIDTH (logical Δy = tableW; the
  // physical box spans x ∈ [cssW − y − tableW, cssW − y]). `w:jc`/`w:tblInd`
  // placement along the flow axis and row-splitting across vertical pages are
  // un-adjudicated follow-ups; the paginator charges the same tableW footprint
  // (the canonical upright-table branch) so pagination and paint agree.
  //
  // WON'T-FIX (narrowed, issue #988 re-adjudication 2026-07-12; measurements in
  // docx-vertical-table-cells.probe.test.ts): Word additionally flows the
  // vertical text BELOW each top-anchored table and places a MULTI-table run in
  // a page-level order that is NON-CAUSAL for a single forward pass — a later
  // table displaces earlier text, tables progress left→right (reverse of the RTL
  // text), and an auto table can overhang the right margin. §17.6.20/§17.4.80 do
  // not define this and a plain block table has no wrapTopAndBottom, so we keep
  // the causal RTL block-flow placement here rather than sample-fit a Word quirk
  // (independently reviewed). Reproducing it would need a place-all-tables →
  // register-exclusions → reflow pass, gated on a purpose-built fixture matrix.
  if (state.verticalPhys) {
    const cssW = state.verticalPhys.cssWidthPx;
    const physState = verticalPhysicalContentState(state);
    const { colWidths, tableW, rowHeights } = computeTableLayout(
      table, physState.contentW, physState,
    );
    const physX = cssW - state.y - tableW;
    // The column axis start: the LOGICAL band start (state.contentX) images the
    // physical top of the current column (physical y = logical x under the +90°
    // page paint) — the top content margin for a single-column body.
    const physY = state.contentX;
    const { ctx } = state;
    ctx.save();
    ctx.rotate(-Math.PI / 2);
    ctx.translate(-cssW, 0);
    drawTableRows(table, colWidths, tableW, rowHeights, physX, physY, physState);
    ctx.restore();
    state.y += tableW;
    return;
  }

  const { contentX, contentW, scale } = state;

  // ECMA-376 §17.4.50 `<w:tblInd>` — indentation added before the table's LEADING
  // edge, shifting it into the text margin. It applies ONLY when the resolved `jc`
  // is left/leading (§17.4.50: "if the resulting justification … is not left …
  // this property shall be ignored"). A NEGATIVE indent pulls the table OUTWARD
  // past the leading margin toward the page edge (sample-28's header banner). Such
  // a table legitimately extends into the page margins and keeps its full
  // preferred width, so widen the layout budget to the whole page (otherwise
  // `resolveColumnWidths`' content-width fit would scale the banner down to the
  // narrower text column and it would never reach the page edge).
  const applyInd = table.tblInd != null && table.jc === 'left';
  const layoutBudget =
    applyInd && (table.tblInd as number) < 0 ? state.pageWidth * scale : contentW;
  const { colWidths, tableW, rowHeights } = computeTableLayout(table, layoutBudget, state);

  // Horizontal table alignment on the page (w:tblPr/w:jc).
  let tableX =
    table.jc === 'center'
      ? contentX + Math.max(0, (contentW - tableW) / 2)
      : table.jc === 'right'
        ? contentX + Math.max(0, contentW - tableW)
        : contentX;

  if (applyInd) {
    // §17.4.50 places the table's LEADING edge `tblInd` inward from the leading
    // text margin (so a NEGATIVE indent pushes it OUTWARD into the margin).
    // `drawTableRows` always takes the physical LEFT origin (`tableX`) and mirrors
    // the columns internally for RTL, so resolve the leading edge to a left origin:
    //   • LTR — leading edge = LEFT text margin (contentX). Left origin =
    //     contentX + tblInd.
    //   • RTL (`bidiVisual`) — leading edge = RIGHT text margin
    //     (contentX + contentW). Its RIGHT edge sits `tblInd` inward from there,
    //     i.e. rightEdge = contentX + contentW − tblInd, so the left origin is
    //     rightEdge − tableW.
    const indPx = (table.tblInd as number) * scale;
    tableX =
      table.bidiVisual === true
        ? contentX + contentW - indPx - tableW
        : contentX + indPx;
  }

  const y = drawTableRows(table, colWidths, tableW, rowHeights, tableX, state.y, state);
  state.y = y;
}

/** Height (px) of a single table row via the shared ST_HeightRule skeleton
 *  ({@link resolveSingleRowHeight}), with the paint pass's px cell measurer.
 *  EXCLUDES the §17.4.85 vMerge span extension — `computeTableLayout` applies
 *  that across the whole table. Exported for unit tests (table-row-height.test). */
export function calculateRowHeight(
  row: DocTableRow,
  table: DocTable,
  colWidths: number[],
  scale: number,
  state: RenderState,
): number {
  return resolveSingleRowHeight(row, colWidths, scale, (cell, cellW) =>
    measureCellContentHeightPx(cell, table, cellW, scale, state),
  );
}

function measureCellParagraphHeight(
  state: RenderState,
  para: DocParagraph,
  maxWidth: number,
  scale: number,
): number {
  return measureCellParagraphWindow(state, para, maxWidth, scale).heightPx;
}

/** Slice-aware twin of {@link measureCellParagraphHeight}: measures only the
 *  `[range.start, range.end)` line window (a mid-row split piece's slice) with
 *  the SAME canonical/legacy scale bridge as paint, and reports the paragraph's total
 *  line count so the caller can tell whether the window covers the paragraph
 *  end (trailing-spacing ownership). No range ⇒ the full paragraph — byte-
 *  identical to the historical behavior. The invariant is that vAlign and row
 *  sizing consume the same line boxes that paint actually draws. */
function measureCellParagraphWindow(
  state: RenderState,
  para: DocParagraph,
  maxWidth: number,
  scale: number,
  range?: { start: number; end: number },
): { heightPx: number; totalLines: number } {
  {
    const paragraphContext = resolveStateParagraphLayoutContext(state, para);
    const grid = gridForParagraphContext(state, paragraphContext);
    const availableWidthPt = maxWidth / scale;
    const measured = measureParagraph(
      para,
      paragraphContext,
      {
        startYPt: 0,
        paragraphXPt: 0,
        availableWidthPt,
        maximumYPt: state.pageH / scale,
        suppressSpaceBefore: true,
      },
      {
        context: state.ctx,
        fontFamilyClasses: state.fontFamilyClasses,
      },
      paragraphMeasurementEnvironment(state),
    );
    // measureParagraph works in scale-1 points (its contract). Reproduce the
    // SAME scale bridge as renderParagraph: ordinary body-table text keeps these
    // canonical line boxes and paint maps them through the Canvas viewport;
    // excluded legacy paths re-measure their line geometry at paint scale. This
    // keeps vAlign and content-driven row sizing equal to the painted glyph box.
    const scale1ContentHeight = measured.contentEndYPt - measured.placement.startYPt;
    const totalLines = measured.markOnly ? 0 : measured.lines.length;
    const windowStart = range ? Math.max(0, range.start) : 0;
    const windowEnd = range ? Math.min(totalLines, range.end) : totalLines;
    if (scale === 1) {
      if (!range || measured.markOnly || measured.lines.length === 0) {
        return { heightPx: scale1ContentHeight, totalLines };
      }
      // Windowed scale-1 content height: the same per-line extents the
      // paginator charges (advance + any non-negative gap to the previous
      // line's bottom; cells carry no wrap oracle, so gaps are zero).
      let sum = 0;
      for (let i = windowStart; i < windowEnd; i++) {
        const line = measured.lines[i];
        if (i === windowStart) { sum += line.advancePt; continue; }
        const previous = measured.lines[i - 1];
        sum += Math.max(0, line.topYPt - (previous.topYPt + previous.advancePt)) + line.advancePt;
      }
      return { heightPx: sum, totalLines };
    }
    const paraHasRuby = paragraphContext.hasRuby;
    const eastAsian = paragraphContext.hasEastAsianText;
    if (measured.markOnly || measured.lines.length === 0) {
      // Empty / anchor-only paragraph mark (§17.3.1.29): renderEmptyMarkParagraph
      // reserves the mark-line height at the PAINT scale, not scale-1 × scale.
      return {
        heightPx: paragraphMarkLineHeight(
          para,
          scale,
          grid,
          paraHasRuby,
          state.docEastAsian,
          state.ctx,
          state.fontFamilyClasses,
          paragraphContext.lineSpacing,
          state.resolvedLocalFonts,
          state.layoutServices?.text,
          paragraphMarkShapeInput(para),
        ),
        totalLines,
      };
    }
    const segments = buildSegments(para.runs, segmentEnvironmentOf(state));
    const canonicalTextScale = canonicalParagraphTextScaleEligible(
      state.storyContext ?? BODY_STORY_CONTEXT,
      state.verticalCJK,
      false,
      false,
      paragraphContext,
      para,
      segments,
    );
    // Rehydrate the scale-1 line PARTITION exactly as renderParagraph does:
    // canonical body-table text scales stored geometry, while excluded legacy
    // paths re-measure it. Then advance by the per-line box height with the SAME
    // ruby/docGrid/lineSpacing resolver
    // (§17.3.1.33). A table cell carries no page-level float wrap oracle, so no
    // line has a topY jump; the painted content height is Σ lineHForLine over the
    // whole, unsliced paragraph.
    const paintLines = rescaleLayoutLines(
      measured.lines.map((line) => line.layout),
      scale,
      state.ctx,
      state.fontFamilyClasses,
      gridCharDeltaPx(grid, scale),
      canonicalTextScale,
    );
    const uniformLineH = paraHasRuby
      ? snapParaLineToGrid(
          Math.max(0, ...paintLines.map((l) => lineBoxHeight(
            para.lineSpacing, l.ascent, l.descent, scale, grid, true, l.intendedSingle, eastAsian,
          ))),
          grid,
          scale,
        )
      : 0;
    const lineHForLine = (l: LayoutLine): number =>
      paraHasRuby
        ? uniformLineH
        // §17.6.5 cell rounding is gated by the line's script; a Latin-only line
        // in a CJK paragraph keeps its natural height, matching the text-box path.
        : lineBoxHeight(para.lineSpacing, l.ascent, l.descent, scale, grid, false, l.intendedSingle, l.eastAsian ?? false, l.gridCountSingle);
    return {
      heightPx: paintedParagraphHeight(paintLines, windowStart, windowEnd, 0, lineHForLine),
      totalLines,
    };
  }
}

/** Effective cell margins (pt). Per-cell `<w:tcMar>` overrides (ECMA-376
 *  §17.4.42) take precedence per edge over the table-level `<w:tblCellMar>`
 *  default (§17.4.41). A résumé template, for example, gives one cell a larger
 *  top margin to add space above its content. */
function effCellMargins(
  cell: DocTableCell,
  table: DocTable,
): { top: number; bottom: number; left: number; right: number } {
  return {
    top: cell.marginTop ?? table.cellMarginTop,
    bottom: cell.marginBottom ?? table.cellMarginBottom,
    left: cell.marginLeft ?? table.cellMarginLeft,
    right: cell.marginRight ?? table.cellMarginRight,
  };
}

function measureCellContent(
  cell: DocTableCell,
  table: DocTable,
  cellW: number,
  scale: number,
  state: RenderState,
): void {
  const cellState = withTableCellStory(state);
  const cm = effCellMargins(cell, table);
  const ml = cm.left * scale;
  const mr = cm.right * scale;
  const innerW = cellW - ml - mr;
  for (const ce of cell.content) {
    measureCellElementHeight(cellState, ce, innerW, scale);
  }
}

/** Measure a cell-level element (paragraph or nested table) at the rendering
 *  scale. Returns total occupied height including paragraph spacing. */
function measureCellElementHeight(
  state: RenderState,
  ce: CellElement,
  innerWPx: number,
  scale: number,
): number {
  if (ce.type === 'paragraph') {
    const para = ce as unknown as DocParagraph;
    // §17.3.1.7: the paint pass (renderCellContent → renderParagraph) advances
    // `max(spaceAfter, bottomBorderExtentPt)` below the text box so following
    // content clears a drawn bottom border. Mirror it here, or a bordered cell
    // paragraph paints taller than the cell measures (B2: single measurer).
    // renderCellContent never passes a borderMerge, so no suppression term.
    //
    // Mid-row split pieces: a sliced cell paragraph (runtime `lineSlice` on the
    // piece clone) measures ONLY its window — leading spacing belongs to the
    // slice that starts the paragraph, trailing spacing to the slice that ends
    // it (the split walk charges them the same way). An unsliced element is
    // byte-identical to the historical measure.
    const slice = (ce as CellElement & { lineSlice?: { start: number; end: number } }).lineSlice;
    const { heightPx, totalLines } = measureCellParagraphWindow(state, para, innerWPx, scale, slice);
    const leading = !slice || slice.start === 0 ? para.spaceBefore : 0;
    const trailing = !slice || slice.end >= totalLines
      ? Math.max(para.spaceAfter, bottomBorderExtentPt(para.borders))
      : 0;
    return heightPx + (leading + trailing) * scale;
  }
  // Nested table — estimateTableHeight works in pt; convert to px.
  const tbl = ce as unknown as DocTable;
  return estimateTableHeight(state, tbl, innerWPx / scale) * scale;
}

function renderCell(
  cell: DocTableCell,
  table: DocTable,
  x: number,
  y: number,
  w: number,
  h: number,
  state: RenderState,
  clipExact = false,
): void {
  const { ctx, scale } = state;

  // Cell BACKGROUND + content only. Borders are painted in a separate, later
  // pass by drawTableRows (ECMA-376 §17.4.66 border-collapse): a shared gridline
  // must sit on top of every cell fill, so no neighbouring cell's background can
  // occlude the border drawn by the cell on the other side of the gridline.
  if (cell.background) {
    ctx.fillStyle = `#${cell.background}`;
    ctx.fillRect(x, y, w, h);
  }

  const cm = effCellMargins(cell, table);
  const mt = cm.top * scale;
  const mb = cm.bottom * scale;
  const ml = cm.left * scale;
  const mr = cm.right * scale;

  // ECMA-376 §17.6.5 defines w:docGrid as a section-level constraint on
  // Cell paragraphs inherit the section's docGrid, but their line-spacing
  // rule comes from the table style's pPr (see parse_table + StyleMap's
  // `resolve_para` with a `table_style_id`). "Table Grid" sets line=240
  // (M=1.0), so with docGrid a cell line box is `max(natural, pitch × 1.0)`
  // = pitch (~18pt), matching Word's observed in-cell baseline advance
  // on demo/sample-1 page 3.
  const cellState: RenderState = {
    ...state,
    contentX: x + ml,
    contentW: w - ml - mr,
    y: y + mt,
    storyContext: enterTableCellStoryContext(
      state.storyContext ?? BODY_STORY_CONTEXT,
    ),
    // ECMA-376 §17.3.2.6 — expose the cell fill (§17.4.33 `<w:tcPr><w:shd>`) as the
    // effective background so an automatic run color inside the cell contrasts
    // against it (sample-28 p.17: a near-black `w:fill="0C0C0C"` cell flips its
    // color-less text to white). A cell with no fill inherits any outer container
    // background (e.g. a nested table). renderParagraph narrows this to the
    // paragraph shading when the paragraph declares its own.
    containerShading: cell.background ?? state.containerShading,
    // ECMA-376 §17.4.57 / §20.4.2.x — a table cell is its own text container: the
    // page's floating objects (anchor images, text frames, and floating TABLES)
    // exclude MAIN-STORY text, NOT text inside a table cell. Word never flows cell
    // content around a page float that happens to overlap the cell's box. Spreading
    // `state.floats` into the cell made a cell paragraph's line layout skip past an
    // OUTER float's wrap band (skipPastTopAndBottom / resolveLineFloatWindow read
    // `state.floats`), pushing the cell's first line down — measured on sample-28
    // p.17, where a vAlign="center" header cell's text was displaced ~17 px below
    // its centred slot by the projects float's band overlapping the cell. Give the
    // cell an isolated (empty) float set so its content lays out only against the
    // cell box; an in-cell anchor float then also stays scoped to the cell instead
    // of leaking onto the page. floatParaSeq restarts at 0 for the same isolation.
    floats: [],
    floatParaSeq: 0,
  };

  if (cell.vAlign === 'center' || cell.vAlign === 'bottom') {
    // ECMA-376 §17.4.7 requires every <w:tc> to end with a <w:p>. When a cell's
    // visible content is a nested table, Word emits a trailing empty paragraph
    // purely as that syntactic anchor. Including it in the centering content
    // height would balloon contentH ≈ rowHeight and pin the visible block to
    // the top of the cell — matching neither Word nor LibreOffice's
    // rendering of resume "bar chart" cells. Skip a single trailing empty
    // paragraph after a non-paragraph block.
    const visibleContent = trimTrailingStructuralMarker(cell.content);
    // ONE vAlign content authority for split and unsplit cells alike: the
    // slice-aware, real-scale measure (measureCellElementHeight honors a piece
    // clone's `lineSlice`, so a mid-row piece centres its OWN window — the
    // Finding-1 invariant keeps the rescale machinery, never a scale-1 sum
    // × scale). The box is the DRAWN cell box `h` (for a vMerge restart piece
    // that is the span box on this page, which is exactly what Word centres
    // against — measured on the split-form ground truth).
    // ECMA-376 §17.3.1.33 + §17.4.83 (vAlign): Word collapses the FIRST
    // paragraph's space-before and the LAST paragraph's space-after against the
    // cell's content boundary when vertically aligning. Neither produces any ink
    // (nothing surrounds them inside the cell), so including them in the
    // vertically-aligned block height pushes the visible block off centre/bottom.
    // Word vertically aligns the INKED block alone: a header cell whose only
    // paragraph carries 6 pt space-before + 8 pt space-after still centres the
    // ~16.8 pt line box, not 30.8 pt. The symmetric trim mirrors how block
    // spacing collapses at a container edge (§17.3.1.33 describes spacing
    // BETWEEN paragraphs, not at the frame boundary). Spacing BETWEEN two
    // paragraphs inside the cell is left intact (handled by §17.3.1.33's
    // contextual / max-overlap rules inside the paint pass).
    const firstEl = visibleContent[0];
    const sliceOf = (el: CellElement | undefined) =>
      (el as (CellElement & { lineSlice?: { start: number; end: number } }) | undefined)?.lineSlice;
    let contentH = 0;
    let previousParagraph: DocParagraph | null = null;
    let previousAfterPx = 0;
    let haveVisibleBlock = false;
    for (const element of visibleContent) {
      if (element.type === 'table') {
        if (previousParagraph) contentH += previousAfterPx;
        contentH += measureCellElementHeight(cellState, element, w - ml - mr, scale);
        previousParagraph = null;
        previousAfterPx = 0;
        haveVisibleBlock = true;
        continue;
      }
      const paragraph: DocParagraph = element;
      const slice = sliceOf(element);
      const window = measureCellParagraphWindow(
        cellState,
        paragraph,
        w - ml - mr,
        scale,
        slice,
      );
      const ownsBefore = !slice || slice.start === 0;
      const ownsAfter = !slice || slice.end >= window.totalLines;
      const beforePx = ownsBefore ? paragraph.spaceBefore * scale : 0;
      const afterPx = ownsAfter ? paragraph.spaceAfter * scale : 0;
      const lineBlockPx = window.heightPx;
      const gapPx = previousParagraph
        ? paragraphGapPt(
            previousParagraph,
            paragraph,
            previousAfterPx / scale,
            beforePx / scale,
          ) * scale
        : haveVisibleBlock ? beforePx : 0;
      contentH += gapPx + lineBlockPx;
      previousParagraph = paragraph;
      previousAfterPx = afterPx;
      haveVisibleBlock = true;
    }
    // Leading space-before (first paragraph only). Nested table first ⇒ 0.
    // A continuation slice (start > 0) charged no space-before in the measure,
    // so there is nothing to trim (and renderParagraph suppresses it too).
    const firstSlice = sliceOf(firstEl);
    const firstSpaceBefore = firstEl && firstEl.type === 'paragraph' && (!firstSlice || firstSlice.start === 0)
        ? (firstEl as unknown as DocParagraph).spaceBefore * scale
        : 0;
    // `renderParagraph` will re-consume the first paragraph's spaceBefore (it
    // unconditionally adds `para.spaceBefore * scale` to `state.y`). Pull
    // `cellState.y` up by `firstSpaceBefore` so that addition lands the inked
    // top exactly on the vertically-aligned position. Without this pull-up the
    // visible block lands `firstSpaceBefore` PAST the intended vAlign position
    // (= +3 pt down for a typical 6 pt spaceBefore at scale 1) — asymmetric with
    // the trailing-spaceAfter trim, which renderParagraph never reconsumes
    // because nothing follows it inside the cell.
    if (cell.vAlign === 'center') {
      cellState.y = y + (h - contentH) / 2 - firstSpaceBefore;
    } else {
      cellState.y = y + h - contentH - mb - firstSpaceBefore;
    }
  }

  if (clipExact) {
    // ECMA-376 §17.4.80 (trHeight) + §17.18.37 (ST_HeightRule "exact"):
    // the row height is exactly @val and content taller than that must not
    // bleed into adjacent rows. The clip is therefore **Y-axis only** — the
    // spec puts no horizontal bound on cell content. Clipping the full
    // (x, y, w, h) bbox half-masks a 0.5 pt nested-table border that lands
    // exactly on the cell's left/right edge (e.g. outer tcMar.left=0 +
    // tblCellMar.left=0 + inner tblInd=0): half the stroke straddles the clip
    // boundary and visibly disappears. Clipping by Y alone preserves the
    // anti-bleed intent without erasing borders that legitimately sit on the
    // cell edge.
    ctx.save();
    ctx.beginPath();
    // `ctx.canvas.width` is the PHYSICAL device width; on a vertical (§17.6.20
    // tbRl) page the ctx is rotated, so this clip rect is expressed in LOGICAL
    // coordinates and the physical-width span makes it OVER-wide along the logical
    // x-axis. That is harmless here — the clip is a Y-band anti-bleed guard, so an
    // over-wide x-extent still fully contains the intended row band (it never
    // UNDER-clips). Vertical tables are not yet exercised by a ground-truth
    // fixture; when they are, tighten this to the logical content width.
    ctx.rect(0, y, ctx.canvas.width, h);
    ctx.clip();
    renderCellContent(cell.content, cellState);
    ctx.restore();
  } else {
    renderCellContent(cell.content, cellState);
  }
}

/** Drop a trailing empty paragraph that follows a non-paragraph block (nested
 *  table). ECMA-376 §17.4.7 requires every cell to end with a paragraph; when
 *  the visible content is a nested table, Word's emitted trailing <w:p/> is a
 *  structural anchor with no visible role. Returns the original array if no
 *  such pattern matches. */
function trimTrailingStructuralMarker(content: CellElement[]): CellElement[] {
  if (content.length < 2) return content;
  const last = content[content.length - 1];
  const prev = content[content.length - 2];
  if (last.type !== 'paragraph' || prev.type === 'paragraph') return content;
  const lastPara = last as unknown as DocParagraph;
  if (lastPara.runs.length > 0) return content;
  return content.slice(0, -1);
}

/** Render a cell's interleaved paragraphs and nested tables in document order.
 *  Mirrors renderBodyElements but without page-break handling (cells never
 *  contain page breaks in our model). */
function renderCellContent(content: CellElement[], state: RenderState): void {
  let prevPara: DocParagraph | null = null;
  let prevSpaceAfter = 0;
  for (const ce of content) {
    if (ce.type === 'paragraph') {
      const para = ce as unknown as DocParagraph;
      const slice = (ce as CellElement & {
        lineSlice?: { start: number; end: number; continues?: boolean };
      }).lineSlice;
      const continues = slice?.start != null && slice.start > 0;
      const lineSlice = slice
        ? { ...slice, ...(continues ? { continues: true as const } : {}) }
        : undefined;
      // §17.3.1.9 per-side contextualSpacing (contextualSpacingAdjust) over the
      // §17.3.1.33 max-collapse — a continuation slice consumed its spacing
      // boundary on the previous slice (no adjacency, prev = null).
      const adjust = contextualSpacingAdjust(
        continues ? null : prevPara, para, prevSpaceAfter, continues ? 0 : para.spaceBefore);
      state.y -= adjust.overlap * state.scale;
      renderParagraph(para, state, adjust.suppressBefore || continues, lineSlice);
      prevPara = para;
      prevSpaceAfter = para.spaceAfter;
    } else if (ce.type === 'table') {
      renderTable(ce as unknown as DocTable, state);
      prevPara = null;
      prevSpaceAfter = 0;
    }
  }
}

/** Resolve a `nil`/`none` border to "no ink". A `null` means "not set" — the
 *  caller already substituted a fallback before reaching here. */
function paintable(b: BorderSpec | null): BorderSpec | null {
  if (!b) return null;
  if (b.style === 'none' || b.style === 'nil') return null;
  return b;
}

/** Stroke one crisp axis-aligned segment. `perp` shifts the whole line
 *  perpendicular to its direction (px, pre-crisp-snap) — used to place the two
 *  rails of a `double` border on either side of the nominal edge. */
function strokeCrispSegment(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  x1: number, y1: number, x2: number, y2: number,
  lw: number,
  dpr: number,
  perp: number,
): void {
  ctx.lineWidth = lw;
  // Crispness nudge (see crispOffset): a thin (odd device-width) axis-aligned
  // stroke straddles two device rows and blurs; nudging it perpendicular to the
  // line snaps it onto the nearest crisp device position. Cell / paragraph
  // borders are always horizontal (y1===y2) or vertical (x1===x2) — never
  // diagonal — so the orientation is read directly from the endpoints, and the
  // snap delta is derived from the line's own coordinate (fractional-safe).
  const horizontal = y1 === y2;
  const vertical = x1 === x2;
  // `perp` runs along x for a horizontal line, along y for a vertical line.
  const ox = (horizontal ? 0 : perp);
  const oy = (horizontal ? perp : 0);
  const dpx = vertical ? crispOffset(x1 + ox, lw, dpr) : 0;
  const dpy = horizontal ? crispOffset(y1 + oy, lw, dpr) : 0;
  ctx.beginPath();
  ctx.moveTo(x1 + ox + dpx, y1 + oy + dpy);
  ctx.lineTo(x2 + ox + dpx, y2 + oy + dpy);
  ctx.stroke();
}

/**
 * ECMA-376 §17.18.2 ST_Border dash/dot families → a `setLineDash` pattern,
 * expressed in units of the stroked width `lw` (px). Thin wrapper over core's
 * shared `docxBorderDashArray` (which owns the §17.18.2 relative table); the ctx
 * is already `scale(dpr,dpr)`d, so `lw`-relative lengths render crisply at any
 * dpr (matching the single/double paths). Returns `[]` for solid styles.
 * Re-exported here so the existing `border-dash.test.ts` contract is preserved.
 */
export function borderDashPattern(style: string, lw: number): number[] {
  return docxBorderDashArray(style, lw);
}

function drawBorderLine(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  x1: number, y1: number, x2: number, y2: number,
  spec: BorderSpec,
  scale: number,
  dpr = 1,
): void {
  ctx.save();
  ctx.strokeStyle = spec.color ? `#${spec.color}` : '#000000';
  const lw = Math.max(0.5, spec.width * scale);

  if (spec.style === 'double') {
    // ECMA-376 §17.18.2 ST_Border "double": two parallel lines with a gap,
    // painted as device-pixel-aligned rail/gap/rail fills so a thin double
    // (e.g. sz6 ≈ 0.75px) never collapses into one line. Shared with the other
    // renderers via core's `fillDoubleBorder` (see core/draw/double-border.ts).
    ctx.fillStyle = ctx.strokeStyle;
    fillDoubleBorder(ctx, x1, y1, x2, y2, lw, dpr);
    ctx.restore();
    return;
  }

  // Dashed/dotted ST_Border families (§17.18.2). setLineDash is reset by the
  // ctx.restore() below. Solid styles get an empty pattern → continuous line.
  const dash = borderDashPattern(spec.style, lw);
  if (dash.length) ctx.setLineDash(dash);
  strokeCrispSegment(ctx, x1, y1, x2, y2, lw, dpr, 0);
  ctx.restore();
}

/**
 * ECMA-376 §17.3.1.7 — paragraph-border merge context for a run of consecutive
 * identically-bordered paragraphs. Word draws ONE box around such a run:
 *   - the `top` edge only on the FIRST paragraph,
 *   - the `bottom` edge only on the LAST paragraph,
 *   - the `<w:between>` edge (if any) at every INNER join,
 *   - `left`/`right` always (they form the box sides).
 * The paint loops (renderBodyElements / renderParaList) detect adjacency and
 * pass `suppressTop` when a same-border paragraph precedes this one, and
 * `suppressBottom` when one follows. When `suppressTop` is set the `between`
 * edge (if defined) is drawn at the top join instead of the `top` edge.
 */
export interface ParaBorderMerge {
  /** A same-border paragraph is adjacent above ⇒ don't draw this `top` edge
   *  (draw `between` at the top join instead, when defined). */
  suppressTop?: boolean;
  /** A same-border paragraph is adjacent below ⇒ don't draw this `bottom` edge. */
  suppressBottom?: boolean;
}

/** The exact PAINTED height (px) of the lines a {@link renderParagraph} draw pass
 *  puts on this page, computed by replaying the per-line advancement the draw loop
 *  performs — WITHOUT drawing. The shading rect must match the paragraph border's
 *  height, and the border height is `state.y − textAreaTopY` measured AFTER the
 *  loop; but shading is the BACKGROUND and must be filled BEFORE the loop (text
 *  paints on top). So we cannot read the post-loop `state.y` for the fill — we
 *  re-derive it from the same inputs the loop uses.
 *
 *  The loop, for each line `li` in `[sliceStart, paintEnd)`, does exactly:
 *    if (line.topY !== undefined && line.topY > y) y = line.topY;  // float clearance
 *    y += lineHForLine(line);                                       // line box advance
 *  starting from `y = textAreaTopY`. Replaying it here yields H === the loop's final
 *  `state.y − textAreaTopY` BY CONSTRUCTION (same height source), so the shading
 *  meets the bottom border in every case:
 *   - normal (no float/slice): H === Σ lineHForLine over all lines (== the old naive
 *     `totalTextH`, so no regression);
 *   - float clearance: a line whose `topY` jumps past the natural flow grows H to
 *     match the border (previously the naive sum stopped short);
 *   - page-sliced paragraph: only `[sliceStart, paintEnd)` is summed, so H no longer
 *     overfills to the full-paragraph height past the slice's bottom border.
 *  `lineHForLine` is the paragraph-scope resolver (ruby/docGrid/lineSpacing) the
 *  loop already uses; passing it as a callback keeps this pure and testable. */
export function paintedParagraphHeight<L extends { topY?: number }>(
  lines: readonly L[],
  sliceStart: number,
  paintEnd: number,
  textAreaTopY: number,
  lineHForLine: (line: L) => number,
): number {
  let y = textAreaTopY;
  for (let li = sliceStart; li < paintEnd; li++) {
    const line = lines[li];
    if (line.topY !== undefined && line.topY > y) y = line.topY;
    y += lineHForLine(line);
  }
  return y - textAreaTopY;
}

/** ECMA-376 §17.3.1.31 — paragraph shading fills the border BOX, not just the
 *  text extent. §17.3.1.31 itself only says the shading sets the paragraph's
 *  background color and is SILENT on border geometry; the fill-to-border is
 *  observed Word behavior: Word fills the border box, and §17.3.1.7 places each
 *  border's `w:space` OUTSIDE the text box (applied by {@link drawParaBorders}), so
 *  the shading reaches those borders. Return the content box grown by each PRESENT
 *  border's space, using the SAME per-edge conditions as drawParaBorders so the
 *  fill meets the border exactly. Without a bordered edge (or no borders at all)
 *  that edge is not extended. (sample-11: a right border with `space=4` left the
 *  gray box detached from its border because the fill stopped `space` short of it.)
 *  Exported for unit testing the per-edge extension. */
export function paraShadingRect(
  x: number, y: number, w: number, h: number,
  borders: ParagraphBorders | null | undefined,
  merge: ParaBorderMerge | undefined,
  scale: number,
): { x: number; y: number; w: number; h: number } {
  if (!borders) return { x, y, w, h };
  const sp = (edge: ParaBorderEdge | null): number =>
    edge && edge.style !== 'none' ? (edge.space ?? 0) * scale : 0;
  const topEdge = merge?.suppressTop ? borders.between : borders.top;
  const l = sp(borders.left);
  const r = sp(borders.right);
  const t = sp(topEdge);
  const b = merge?.suppressBottom ? 0 : sp(borders.bottom);
  return { x: x - l, y: y - t, w: w + l + r, h: h + t + b };
}

export interface ParaBorderSegment {
  side: 'top' | 'bottom' | 'left' | 'right';
  edge: ParaBorderEdge;
  x1: number;
  y1: number;
  x2: number;
  y2: number;
}

/**
 * Horizontal paragraph content extent used by shading and `w:pBdr`.
 *
 * ECMA-376 §17.3.1.24 applies pBdr to the paragraph. A
 * hanging first line (including a list marker from §17.9) is part of that
 * paragraph, so its start position must be inside the box. The normal body edge
 * remains authoritative for a positive first-line indent because later lines
 * still begin at the body edge. `physicalIndentLeft/Right` have already been
 * mirrored for bidi; `firstIndent` remains logical, so the expansion is applied
 * to the physical left for LTR and physical right for RTL. Resolved marker bounds
 * are then unioned so lvlJc-shifted text and picture markers remain inside the box.
 */
export function paragraphBorderContentBox(
  contentX: number,
  contentW: number,
  physicalIndentLeft: number,
  physicalIndentRight: number,
  firstIndent: number,
  baseRtl: boolean,
  markerBounds?: { left: number; right: number },
): { x: number; w: number } {
  let left = contentX + physicalIndentLeft;
  let right = contentX + contentW - physicalIndentRight;
  if (firstIndent < 0) {
    if (baseRtl) right -= firstIndent;
    else left += firstIndent;
  }
  if (markerBounds) {
    left = Math.min(left, markerBounds.left);
    right = Math.max(right, markerBounds.right);
  }
  return { x: left, w: Math.max(0, right - left) };
}

/**
 * Resolve paragraph-border strokes into one connected box.
 *
 * ECMA-376 §17.3.1.17 requires a left paragraph border to run between the
 * top/between border above and the bottom/between border below. The same
 * geometry applies symmetrically to the right side: horizontal strokes reach
 * the vertical-border positions, and vertical strokes use the horizontal
 * strokes as their endpoints. Computing the four segments together prevents
 * the independent `w:space` offsets from leaving open corners.
 */
export function paraBorderSegments(
  x: number, y: number, w: number, h: number,
  borders: ParagraphBorders,
  merge?: ParaBorderMerge,
  scale = 1,
): ParaBorderSegment[] {
  const visible = (edge: ParaBorderEdge | null): edge is ParaBorderEdge =>
    edge != null && edge.style !== 'none';
  const sp = (edge: ParaBorderEdge | null): number =>
    visible(edge) ? (edge.space ?? 0) * scale : 0;
  // §17.3.1.7 top edge: on a non-first paragraph of a shared run, the `top` edge
  // gives way to the `between` edge drawn at the join (nothing when `between` is
  // absent — the box has no internal rules).
  const topEdge = merge?.suppressTop ? borders.between : borders.top;
  const bottomEdge = merge?.suppressBottom ? null : borders.bottom;
  const leftX = x - sp(borders.left);
  const rightX = x + w + sp(borders.right);
  const topY = y - sp(topEdge);
  const bottomY = y + h + sp(bottomEdge);
  const segments: ParaBorderSegment[] = [];
  if (visible(topEdge)) {
    segments.push({ side: 'top', edge: topEdge, x1: leftX, y1: topY, x2: rightX, y2: topY });
  }
  // The `bottom` edge is skipped entirely when a same-border paragraph follows
  // (the box continues into it; its own join is handled by that paragraph's
  // suppressed-top `between`).
  if (visible(bottomEdge)) {
    segments.push({ side: 'bottom', edge: bottomEdge, x1: leftX, y1: bottomY, x2: rightX, y2: bottomY });
  }
  if (visible(borders.left)) {
    segments.push({ side: 'left', edge: borders.left, x1: leftX, y1: topY, x2: leftX, y2: bottomY });
  }
  if (visible(borders.right)) {
    segments.push({ side: 'right', edge: borders.right, x1: rightX, y1: topY, x2: rightX, y2: bottomY });
  }
  return segments;
}

function drawParaBorders(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
  borders: ParagraphBorders,
  scale: number,
  dpr = 1,
  merge?: ParaBorderMerge,
): void {
  for (const segment of paraBorderSegments(x, y, w, h, borders, merge, scale)) {
    const { edge } = segment;
    const spec: BorderSpec = { width: edge.width, color: edge.color, style: edge.style };
    drawBorderLine(ctx, segment.x1, segment.y1, segment.x2, segment.y2, spec, scale, dpr);
  }
}

/** ECMA-376 §17.3.1.7 — the vertical extent (in scale-1 points) a paragraph's
 *  BOTTOM border adds BELOW the text box, so following content clears it.
 *
 *  §17.3.1.7 places the bottom border `w:space` points below the text ("the space
 *  after the bottom of the text … before this border is drawn"), and §17.3.4 gives
 *  the border its own width (`w:sz`, eighths of a point). {@link drawParaBorders}
 *  strokes the line CENTERED on `textBottom + space`, so its outer (bottom) edge is
 *  at `textBottom + space + width/2`. Word reserves that whole extent in the flow —
 *  a bottom-bordered paragraph pushes the next paragraph BELOW the border rather
 *  than letting the following line box overlap it (the spec is silent on the flow
 *  reservation; this is Word's observed layout, verified against sample-14's
 *  reference-list rule, whose `space=1 sz=12` rule sat ~1.75 pt too high without it).
 *
 *  Returns 0 when there is no visible bottom edge, or when a same-border paragraph
 *  follows (the bottom edge is suppressed by the §17.3.1.7 merge — the box
 *  continues into the next paragraph, so nothing is drawn here to clear). */
function bottomBorderExtentPt(
  borders: ParagraphBorders | null | undefined,
  merge?: ParaBorderMerge,
): number {
  if (!borders || merge?.suppressBottom) return 0;
  const b = borders.bottom;
  if (!b || b.style === 'none') return 0;
  return (b.space ?? 0) + (b.width ?? 0) / 2;
}

// ===== Utilities =====

/** ECMA-376 §17.3.2.4 — two `<w:bdr>` borders belong to the same run-border
 *  group iff their attribute sets are identical. We compare the attributes the
 *  model carries (style/sz/space/color); themeColor/themeTint/shadow/frame are
 *  not modelled, so identical themed borders that differ only in unmodelled
 *  attributes still group (acceptable — the painted frame is identical anyway). */
function runBordersEqual(a: DocxRunBorder, b: DocxRunBorder): boolean {
  return (
    a.style === b.style &&
    a.width === b.width &&
    (a.space ?? 0) === (b.space ?? 0) &&
    (a.color ?? null) === (b.color ?? null)
  );
}

/** Service-less compatibility adapter. It deliberately does not inspect a
 * leading marker scalar; production body/text-box paths use retained per-scalar
 * TextLayoutService spans. */
function markerFontFamily(num: NumberingInfo): string | null {
  return num.fontFamily ?? null;
}

/** Marker glyph as it should be drawn/measured. Symbol/Wingdings markers
 *  (§17.9.x `w:lvlText` + §17.3.2.26 `w:rFonts`) store the glyph as the FONT's
 *  own code point (e.g. Symbol U+F0B7 = "•", Wingdings U+F0A7 = "▪"). Those
 *  private-encoding code points render as tofu in any fallback face, so we
 *  normalize them to the Unicode equivalent up front — keyed on the marker's
 *  requested ascii family, not on the sample. Non-symbol markers (decimals,
 *  roman, CJK bullets) pass through unchanged. */
function markerDisplayText(num: NumberingInfo): string {
  return symbolFontToUnicode(num.text, num.fontFamily ?? null);
}

/** Minimum clear side-gap (px) an EMPTY paragraph-mark line needs before it may
 *  START beside a float rather than flow below the float band — the pilcrow's own
 *  em width (the paragraph-mark font size × scale). Distinct from the 1-inch
 *  CONTENT-line rule (`wordMinLineStartPx`, issue #676): Word keeps an empty mark
 *  beside a float whenever the gap can hold the pilcrow, and drops it below only
 *  when the gap is narrower than that — i.e. effectively a full-width band. See
 *  WORD_MIN_LINE_START_PT's SCOPE note. Grounded from sample-9 p.4 (a full-width
 *  float band → the mark drops below, carrying its wrapNone anchor image, PR
 *  b897bbf) AND sample-12 p.2 (a ~62pt side-gap under 1 inch where the figure's
 *  nine trailing blank-line marks stay beside the float; flowing them below at
 *  1 inch pushed the caption + CONCLUSION onto page 3 — the regression #676
 *  introduced, which this restores). Single source of truth for the literally-empty /
 *  anchor-only paragraph sites — the paint pass `resolveEmptyMarkTop` and the
 *  paginator mirror `flowMarkLine` — so the two agree bit-for-bit. (A content
 *  paragraph's trailing-break empty final line stays on the 1-inch content-line
 *  rule inside `layoutLines`; see WORD_MIN_LINE_START_PT's SCOPE note.) */
function paragraphMarkEmPx(para: DocParagraph, scale: number): number {
  return getDefaultFontSize(para) * scale;
}

// ───────────────────────────────────────────────────────────────────────────
// ECMA-376 §17.6.5 docGrid CHARACTER grid (字詰め). When the section's docGrid
// `type` is "linesAndChars" or "snapToChars" AND a `charSpace` is declared,
// every full-width East-Asian glyph gains a fixed per-EA-glyph spacing delta
//   Δpt = charSpace / 4096   in FLAT POINTS (NEGATIVE = tighter)
// that is INDEPENDENT of font size — it is added to the glyph's MEASURED advance
// (≈1em for full-width EA glyphs), NOT scaled by it. (`gridCharDeltaPx` returns
// exactly `charSpacePt * scale` = charSpace/4096 pt in px; it does not multiply
// by the font size.) Latin / digits are NOT snapped (they keep their natural
// advance), so the grid delta applies only to EA code points.
//
// ── The single advance model (measure == draw) ──────────────────────────────
// To make line-break MEASUREMENT and the draw ADVANCE provably identical, the
// grid delta enters in exactly ONE way: as a per-code-point spacing on a
// PURE-EA segment. `gridSegDeltaPx` returns the total delta a segment's box
// gains (`len × Δpx` for a pure-EA segment, else 0 — mixed/Latin segments get
// no grid effect, sidestepping any contextual-metric or justification drift),
// and `segAdvanceWidth` folds it into the run's complete advance together with
// §17.3.2.43 `w:w` and §17.3.2.35 `w:spacing`. BOTH the layout's `measuredWidth`
// and every draw path derive the segment's advance from this SAME quantity:
//   • non-justified draw walks the glyphs via `justifiedPiecePositions(cps,
//     [1..n-1], perGap=0, measure, letterSpacingPx=Δ)`, whose final glyph lands
//     at `measure(whole) + n·Δ` = the box edge;
//   • justified draw reuses the EXISTING `justifiedPiecePositions` path with the
//     same `letterSpacingPx = Δ`, so its box edge is `measure(whole) + n·Δ +
//     nGaps·perGap` = `measuredWidth + internalStretch`.
// Because both come from `measure(prefix) + (cps before)·Δ`, draw never diverges
// from `measuredWidth` by construction — there is no separate per-glyph sum to
// drift against the whole-string measure (約物半角 contextual collapse stays
// honoured). See packages/core/src/text/justify-positions.ts.
