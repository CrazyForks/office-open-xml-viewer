import type {
  DocxDocumentModel, BodyElement, DocParagraph, DocTable, DocTableCell, CellElement,
  DocRun, DocxTextRun, ImageRun, ChartRun, ShapeRun, ShapeText, HeaderFooter,
  HeadersFooters, BorderSpec, ParagraphBorders, ParaBorderEdge, SectionProps,
  PageBorders, PageBorderEdge, DocNote,
} from './types';
import { docxRenderedFontFamilies } from './document-content.js';
import {
  getCachedSvgImageByPath,
  preferVectorBlip,
  mathToMathML,
  recolorSvg,
  crispOffset,
  PT_TO_PX,
  isHTMLCanvas,
  defaultDpr,
  clampCanvasSize,
  DEFAULT_KINSOKU_RULES,
  getCachedBitmapByPath,
  acquireBitmapCacheLease,
  deferBitmapCloseWhileLeased,
  applyDuotone,
  imageNaturalSize,
  drawImageCropped,
  metafileRasterSize,
  docxBorderDashArray,
  fillDoubleBorder,
  renderChart,
} from '@silurus/ooxml-core';
import type {
  MathRenderer,
  KinsokuRules,
  HyperlinkTarget,
  NumberFormat,
  Duotone,
  ResolvedLocalFontMetric,
} from '@silurus/ooxml-core';
import {
  segmentsHaveRtl,
  jcIsFullyJustified,
  jcStretchesLastLine,
} from './bidi-line.js';
import {
  type FloatRect,
  FLOAT_OVERLAP_EPS,
  isWrapFloat,
} from './float-layout.js';
import {
  type KashidaLevel,
} from './kashida-justify.js';
import {
  type FrameBox,
  computeFrameBox,
  frameXContainer,
  pushFloatRect,
} from './frame-geometry.js';
import {
  resolveFloatingTableBoxPt,
} from './float-table-geometry.js';
import {
  xContainer,
  yContainer,
  resolveAnchorX,
  resolveAnchorY,
} from './anchor-geometry.js';
import {
  applyTableRowBoundaryFootprints,
  resolveTableRowContentHeights,
} from './table-geometry.js';
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
import { canvasFontString } from '@silurus/ooxml-core';
import type {
  BlockLayoutAlgorithms,
  BodyFlowRegistryDeltaPt,
  BodyFlowRegistrySnapshotPt,
  DrawingMLCollisionRegistrySnapshotPt,
  LayoutServices,
  FloatRegistryEntryPt,
  FloatRegistrySnapshotPt,
  FloatingTablePlacementLayout,
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
import {
  floatRegistryParticipant,
  resolveBlockFlowAdmission,
  resolvePageAnchoredTableDeferral,
} from './layout/floats.js';
import {
  ExactConvergenceError,
  convergeExactState,
} from './layout/convergence.js';
import { LayoutInvariantError } from './layout/diagnostics.js';
import type { LayoutOptions } from './layout/options.js';
import {
  paginatedFlowHasPaginationDependentFields,
} from './layout/pagination-fields.js';
import {
  createCanvasPaintResourcePainter,
  paintLayoutPageContent,
  paintLayoutPage as paintRetainedLayoutPage,
} from './paint/canvas-page.js';
import { canonicalCanvasPaintResourceHandlers } from './paint/canonical-resource-handlers.js';
import { wordPageLevelAnchorY } from './layout/anchor-compatibility.js';
import {
  wordRubyUniformLineHeightPx,
} from './layout/line-compatibility.js';
import { wordDropsTrailingStructuralCellMarker } from './layout/table-compatibility.js';
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
  createLayoutServicesRuntimeView,
  fieldAcquisitionContextOf,
  layoutVariantStoreOf,
  paintResourceRegistryOf,
  privateResourceLookupOf,
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
  createTextLayoutService,
  classifyDocxFontGeneric,
  EAST_ASIAN_RE,
  snapshotLocalMetrics,
  type GlyphMeasureRequest,
} from './layout/text.js';
import { DOCX_GOOGLE_FONTS, docxFontPreloadNames } from './google-fonts.js';
import {
  effectiveTablePositioning,
  internalDocumentModel,
  numberingMarkerShapeInput,
  paragraphAcquisitionInput,
  paragraphMarkShapeInput,
  publicAnchorBridge,
  tableParticipatesInOrdinaryFlow,
} from './parser-model.js';
import {
  normalizeInternalDocumentModel,
  bodySectionIndexInput,
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
} from './layout/numbering-marker.js';
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
  isGridLineRule,
  layoutLines,
  lineBoxHeight,
  normalizeFontFamilyUncached,
  paragraphMarkLineHeight,
  rescaleLayoutLines,
} from './line-layout.js';
import type {
  DocGridCtx,
  LayoutLine,
  LayoutSeg,
} from './line-layout.js';
import {
  measureParagraph,
  type ParagraphMeasurementEnvironment,
} from './paragraph-measure.js';
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
} from './layout/table-pagination.js';
import { paragraphGapAdjustment } from './layout/paragraph-spacing.js';
import { imageResourceKey } from './layout/source-key.js';
import {
  resolveParagraphBorderEdges,
} from './layout/paragraph-border-adjacency.js';
import {
  acquireParagraphResult,
  acquireRetainedFrameGroup,
  bodyFrameGroupFor,
  bodyParagraphBorderEdgesFor,
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
  drawTateChuYokoRun,
  drawUprightBox,
  physicalToLogicalAnchorBox,
} from './vertical-text.js';
import { textRunsForPage } from './text-run-projection.js';

function kashidaLevelOf(alignment: string | null | undefined): KashidaLevel | null {
  if (alignment === 'lowKashida') return 'low';
  if (alignment === 'mediumKashida') return 'medium';
  if (alignment === 'highKashida') return 'high';
  return null;
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
   *  {@link verticalCJK}) on a section-level `btLr` page. Under
   *  `word-section-btlr-tbrl-page-frame`, `btLr` shares the `tbRl` PAGE FRAME (swapped logical
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
   *  (its run-level `<w:shd>`). The layout acquisition state carries the cell
   *  fill and paragraph override; absent means the page background. Only the
   *  auto-contrast decision reads it; retained paint owns the shading geometry. */
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
 *  ST_TextDirection enum (ECMA-376 Part 4 §14.11.7, rather than the Part 1
 *  §17.18.93 Strict `tb|rl|lr|…` set):
 *    - `tbRl`  (≡ Strict `rl`)  — vertical, lines right→left: standard vertical
 *                                 Japanese; the only value in the samples.
 *    - `tbRlV` (≡ Strict `rlV`) — vertical R→L, non-EA glyphs rotated 90° CW.
 *    - `tbLrV` (≡ Strict `lrV`) — vertical L→R, non-EA glyphs rotated 90° CW.
 *  These three share the +90° page rotation + upright-CJK glyph path (stage-1
 *  approximates the `V` variants' non-EA rotation the same as `tbRl`, which the
 *  glyph path already draws Latin sideways for).
 *
 *    - `btLr`  (≡ Strict `lr`)  — its NOMINAL semantics are bottom-to-top /
 *                                 left-to-right, but the registered
 *                                 `word-section-btlr-tbrl-page-frame` behavior
 *                                 makes the section
 *                                 `btLr` uses the horizontal layout rotated +90° CW
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
 *  section-level `btLr` under `word-section-btlr-tbrl-page-frame` (see
 *  {@link RenderState.verticalAllRotated}). The tbRl family keeps
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
 *  LOGICAL geometry back to its PHYSICAL page geometry. Under
 *  `word-vertical-section-physical-header-footer`, its header/footer stay
 *  horizontal at the physical top/bottom margins and do not rotate with the tbRl
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
  if (opts.onTextRun) {
    for (const run of textRunsForPage(retainedBodyLayout, pageIndex, { scale })) {
      opts.onTextRun(run);
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
          tableOverlap: request.overlap,
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
            kind: 'table', tableOverlap: 'never', mode: 'square',
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
              && resolvePageAnchoredTableDeferral({
                bounds: {
                  xPt: raw.x,
                  yPt: raw.y,
                  widthPt: raw.w,
                  heightPt: raw.h,
                },
                blockers: floatRegistry.entries.map(floatRegistryParticipant),
                overlapEpsilonPt: FLOAT_OVERLAP_EPS,
              }).defer;
            if (pageAnchoredCollision) {
              // `word-page-anchored-table-collision-deferral`: a fresh page
              // preserves the authored absolute anchor instead of converting
              // the colliding table to a text continuation.
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
            type FloatingParentTransactionPass =
              | Readonly<{
                  kind: 'fresh-flow-region';
                  result: ReturnType<typeof takeTableFragment>;
                }>
              | Readonly<{
                  kind: 'candidate';
                  parentFrame: Readonly<{ xPt: number; yPt: number }>;
                  result: ReturnType<typeof takeTableFragment>;
                  fragment: NonNullable<ReturnType<typeof takeTableFragment>['fragment']>;
                  resolved: ReturnType<typeof resolveFloatingTablePlacementInTransaction>;
                  nestedEntries: readonly FloatRegistryEntryPt[];
                  fingerprint: string;
                }>;
            let transaction: FloatingParentTransactionPass;
            try {
              transaction = convergeExactState<FloatingParentTransactionPass>({
                step: (previous) => {
                  if (previous?.kind === 'fresh-flow-region') return previous;
                  if (previous?.kind === 'candidate'
                    && previous.resolved.placement.xPt === previous.parentFrame.xPt
                    && previous.resolved.placement.yPt === previous.parentFrame.yPt) {
                    return previous;
                  }
                  const parentFrame = previous?.resolved.placement ?? {
                    xPt: raw.x,
                    yPt: raw.y,
                  };
                  const availableHeightPt = Math.max(
                    0,
                    admissionBlockEndPt - parentFrame.yPt,
                  );
                  const result = takeTableFragment(retained, cursor, {
                    availableHeightPt,
                    freshPageHeightPt: freshAdmissionHeightPt,
                    placement: {
                      container: {
                        id: `${request.location.flowDomainId}:floating-table`,
                        kind: 'body',
                        bounds: {
                          xPt: 0,
                          yPt: 0,
                          widthPt: request.availableInlineExtentPt,
                          heightPt: availableHeightPt,
                        },
                      },
                      cursor: { xPt: 0, yPt: 0 },
                      availableBounds: {
                        xPt: 0,
                        yPt: 0,
                        widthPt: request.availableInlineExtentPt,
                        heightPt: availableHeightPt,
                      },
                    },
                    services,
                    compatibility: 'word',
                    oversizedRowPolicy: 'atomic',
                    page: {
                      physicalPageIndex: request.location.pageIndex,
                      displayPageNumber: state.displayPageNumber
                        ?? request.location.pageIndex + 1,
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
                      kind: 'fresh-flow-region' as const,
                      result,
                    });
                  }
                  const sourcePlacement: FloatingTablePlacementLayout = Object.freeze({
                    kind: 'floating-table-placement',
                    occurrenceId: `${retained.input.id}:root:${request.location.pageIndex}:${cursor.rowIndex}:${cursor.rowFragmentIndex}`,
                    ownership: 'source',
                    physicalPageIndex: request.location.pageIndex,
                    displayPageNumber: state.displayPageNumber
                      ?? request.location.pageIndex + 1,
                    hostCellId: request.location.flowDomainId,
                    sourceBlockIndex: request.input.source.path[0]!,
                    anchorBlockIndex: request.input.source.path[0]!,
                    tableId: result.fragment.id,
                    overlap: table.overlap === 'never' ? 'never' : 'overlap',
                    positioning,
                    anchorBounds: frames.text,
                    child: result.fragment,
                  });
                  const nestedEntries = result.floatingTableRegistryDelta?.entries ?? [];
                  const nestedNextParagraphId =
                    result.floatingTableRegistryDelta?.nextParagraphId
                    ?? floatRegistry.nextParagraphId;
                  const resolved = resolveFloatingTablePlacementInTransaction(
                    sourcePlacement,
                    frames,
                    beginFloatingTablePlacementTransaction(
                      floatRegistry.entries,
                      nestedNextParagraphId,
                      floatRegistry.coordinateSpace,
                      floatRegistry.flowDomainId,
                    ),
                  );
                  const fingerprint = JSON.stringify({
                    parentFrame: {
                      xPt: resolved.placement.xPt,
                      yPt: resolved.placement.yPt,
                    },
                    fragment: result.fragment,
                    nestedEntries,
                    resolvedBounds: resolved.placement.bounds,
                  });
                  return Object.freeze({
                    kind: 'candidate' as const,
                    parentFrame: Object.freeze({
                      xPt: parentFrame.xPt,
                      yPt: parentFrame.yPt,
                    }),
                    result,
                    fragment: result.fragment,
                    resolved,
                    nestedEntries,
                    fingerprint,
                  });
                },
                stateOf: (value) => value.kind === 'fresh-flow-region'
                  ? 'fresh-flow-region'
                  : value.fingerprint,
                limit: 16,
              }).value;
            } catch (error) {
              if (error instanceof ExactConvergenceError) {
                throw new LayoutInvariantError(
                  'NON_CONVERGENCE',
                  error.reason === 'cycle'
                    ? 'Floating table parent/child transaction repeated an exact-state cycle'
                    : 'Floating table parent/child transaction reached the operational pass limit 16',
                );
              }
              throw error;
            }
            if (transaction.kind === 'fresh-flow-region') {
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
            const {
              result,
              fragment,
              resolved,
              nestedEntries,
            } = transaction;
            const isFloatingContinuation = request.cursor?.kind === 'table'
              && request.cursor.floatingContinuationFrame !== undefined;
            const admittedBlockEndPt = request.location.availableBounds.yPt
              + request.location.availableBounds.heightPt;
            const hostFlowPlacements = [
              ...fragment.resolvedFloatingTables ?? [],
              resolved.placement,
            ].filter((placement) => placement.source.positioning.vertAnchor === 'text');
            if (!isFloatingContinuation && hostFlowPlacements.some((placement) => (
              placement.exclusionBounds.yPt + placement.exclusionBounds.heightPt
                > admittedBlockEndPt
            ))) {
              return Object.freeze({
                layout: fragment,
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
              layout: fragment,
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
          const retryAtBlockStartPt = resolveBlockFlowAdmission({
            inlineStartPt: tableInlineStartPt,
            inlineEndPt: tableInlineEndPt,
            blockStartPt: request.location.cursorPt.yPt,
            blockExtentPt: remainingTableExtentPt,
            blockers: floatRegistry.entries.map(floatRegistryParticipant),
            overlapEpsilonPt: FLOAT_OVERLAP_EPS,
          }).blockStartPt;
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
            const core = {
              mode: (entry.wrap === 'topAndBottom'
                ? 'topAndBottom'
                : 'square') as FloatRect['mode'],
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
            return entry.kind === 'table'
              ? {
                  ...core,
                  kind: 'table',
                  tableOverlap: entry.overlap,
                }
              : { ...core, kind: entry.kind };
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

function retainedTableRecord(state: RenderState, sourceIndex: number): RetainedTableRecord {
  const record = state.retainedTablesBySourceIndex?.get(sourceIndex);
  if (!record) throw new Error('Table placement requires retained table acquisition');
  return record;
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

/** The layout bridge's scale-1 table acquisition: resolve column widths once,
 *  acquire the retained table, and return its authoritative row advances. The
 *  fallback remains for reduced test/story states until measurement moves fully
 *  under layout ownership. */
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

/**
 * ECMA-376 §17.3.1.9 `<w:contextualSpacing>` — the registered
 * `word-contextual-spacing-per-side` semantics, identical in body, table cell,
 * and text box.
 *
 * The §17.3.1.33 collapsed inter-paragraph gap decomposes as
 *   gap = prevContrib + currContrib
 *   prevContrib = prev.spaceAfter                          (the collapse base)
 *   currContrib = max(curr.spaceBefore − prev.spaceAfter, 0)   (the excess)
 * summing to max(after, before). A paragraph whose toggle is set AND whose
 * neighbour shares its paragraph style drops ITS OWN contribution only:
 *   - prev toggles → gap = max(before − after, 0)  — matches the spec's worked
 *     example (after=10pt, before=12pt → 2pt);
 *   - curr toggles → gap = after — the previous spaceAfter contribution remains
 *     intact (the spec-literal net-minus-before reading would give 0);
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
 * Sum the heights of a cell's content elements with the same paragraph-spacing
 * collapse used by canonical cell acquisition. Two collapse rules apply:
 *
 *   - ECMA-376 §17.3.1.9 `<w:contextualSpacing>`: a same-style toggling
 *     paragraph drops its OWN contribution to the inter-paragraph gap
 *     (`word-contextual-spacing-per-side`, projected by
 *     {@link contextualSpacingAdjust}).
 *   - Adjacent-paragraph spacing OVERLAP: the gap between two paragraphs is
 *     `max(prevSpaceAfter, currSpaceBefore)`, not their sum. We subtract the
 *     overlap `min(prevSpaceAfter, effBefore)` so a 12pt space-after followed
 *     by a 12pt space-before contributes 12pt of gap, not 24pt.
 *
 * A nested table (CellElement other than paragraph) resets the
 * prev-paragraph context — the next paragraph after a table spaces from a
 * fresh baseline.
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
  // Transitional story-measurement scope: canonical body pagination prepares a
  // §17.3.1.11 group for every frame; reduced header/footer acquisition states
  // use this local measurement fallback.
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

/**
 * Resolve an anchored shape's page-space bounding box {x,y,w,h} (px). Retained
 * drawing acquisition and float registration share this geometry so the
 * exclusion band matches the painted box.
 *
 * Mirrors the renderer's sizing: sizeRelH/sizeRelV (ECMA-376 §20.4.2.18)
 * override the static extent, and a wgp child scales by the group ratio with its
 * within-group offset scaled in step; resolveAnchorX/Y then place the box. `w`/`h`
 * may be 0/negative for degenerate line presets; a wrap shape with no area
 * registers no float.
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
  // share one geometry. Under `word-vertical-section-physical-drawing-layer`,
  // a `paragraph`/`line`-relative positionV anchors from the PHYSICAL TOP of
  // the anchor paragraph's COLUMN. That physical y is the column
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
 * the box origin BEFORE the typed float placement policy displaces it.
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

/** ECMA-376 §17.6.20 + §20.4.3.x — a RenderState view whose page/margin geometry
 *  is the PHYSICAL (un-rotated) page, used to resolve a DrawingML anchor's
 *  `<wp:positionH/V>` against the physical page for a vertical (tbRl) section
 *  under `word-vertical-section-physical-drawing-layer`. Only
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
 *  vertical (tbRl) section: anchored shapes and block tables. Under
 *  `word-vertical-section-physical-drawing-layer`, these resolve and paint
 *  against the un-rotated physical page — cell/label text is
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
  // to the page-relative top edge (Y=0 → inside the top margin). ImageRun
  // carries no pctPos/sizeRel, so those args remain null and the legacy boolean
  // anchorXFromMargin / anchorYFromPara hints still gate page-vs-margin when
  // no raw relativeFrom is present. When align is absent, resolveAnchorX/Y
  // fall back to the offset path.
  if (state.verticalPhys) {
    // `word-vertical-section-physical-drawing-layer`: resolve positionH/V in
    // the physical page frame independently of rotated text flow, resolve the
    // box there, then project it into the swapped logical layout frame. The
    // float-exclusion band and the drawUprightBox-un-swapped painted image
    // therefore share one geometry. A `paragraph`/`line`-relative positionV
    // anchors from the physical top of the anchor paragraph's column. That y is
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

/** `word-page-level-float-prescan` classifies a `<wp:positionV>` reference
 *  that resolves the float's Y INDEPENDENTLY of its anchoring paragraph (vs.
 *  `paragraph` / `line` / `character` which resolve against the paragraph's
 *  top). {@link preRegisterPageFloats} uses this to
 *  hoist such floats to page-start; paragraph-local Y still flows the legacy
 *  per-paragraph path.
 *
 *  An anchor with NO explicit `<wp:positionV>` (anchorYRelativeFrom absent)
 *  still resolves against the page top via the legacy hint
 *  (`anchorYFromPara=false` ⇒ page-absolute offset), so it qualifies as
 *  page-level too. */
function isPageLevelAnchorY(rf: string | null | undefined, fromPara: boolean): boolean {
  return wordPageLevelAnchorY(rf, fromPara);
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

/** Register float exclusions from a paragraph's anchored images, charts, and
 *  shapes so body text wraps around retained drawings
 *  (ECMA-376 §20.4.2.16/.17).
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
  // Implementation-defined heuristic with no ECMA-376 basis:
  // displacing the later document-order float, the "other paragraphs only"
  // gate under allowOverlap=true, and the right-then-down re-seat using dist
  // padding as the float-to-float gap. See layout/floats.ts.
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

/** Reserve the float-exclusion rect for one anchored wrap shape. Retained paint
 *  owns the drawing, so this only pushes an already-represented FloatRect. */
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
  // A degenerate (zero/negative-area) box reserves no exclusion band.
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
    // Retained drawing paint owns the shape; mark the exclusion as already
    // represented so the bitmap resource path skips it.
    drawn: true,
    paraId,
    avoidOverlap: true,
    allowOverlap: true,
  });
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

/** Measures only the `[range.start, range.end)` line window (a mid-row split
 *  piece's slice) through the measurement bridge, and reports the paragraph's total
 *  line count so the caller can tell whether the window covers the paragraph
 *  end (trailing-spacing ownership). No range means the full paragraph. The
 *  invariant is that vAlign, row sizing, and retained paint consume the same
 *  acquired line boxes. */
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
    // measureParagraph works in scale-1 points. Ordinary body-table text keeps
    // these canonical line boxes and retained paint maps them through the
    // viewport; the transitional scale branch below serves reduced contexts.
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
      // Empty / anchor-only paragraph mark (§17.3.1.29) reserves its mark-line
      // height at the requested scale.
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
    // Rehydrate the scale-1 line partition and advance it with the same
    // ruby/docGrid/line-spacing resolver (§17.3.1.33). A table cell carries no
    // page-level float wrap oracle, so its content height is the selected
    // window's accumulated line-box advance.
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
          wordRubyUniformLineHeightPx(
            true,
            paintLines.map((l) => lineBoxHeight(
              para.lineSpacing,
              l.ascent,
              l.descent,
              scale,
              grid,
              true,
              l.intendedSingle,
              eastAsian,
            )),
          ),
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
    // §17.3.1.7: retain `max(spaceAfter, bottomBorderExtentPt)` below the text
    // box so following cell content clears the bottom border. Cell acquisition
    // has no adjacent-paragraph border merge, so no suppression term applies.
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

/** Apply `word-trailing-structural-cell-marker`: drop an empty terminal
 *  paragraph after a non-paragraph cell block. Returns the original array when
 *  no such pattern matches. */
function trimTrailingStructuralMarker(content: CellElement[]): CellElement[] {
  const last = content[content.length - 1];
  const prev = content[content.length - 2];
  const lastParagraphRunCount = last?.type === 'paragraph'
    ? (last as unknown as DocParagraph).runs.length
    : undefined;
  return wordDropsTrailingStructuralCellMarker({
    contentLength: content.length,
    previousKind: prev?.type,
    lastKind: last?.type,
    lastParagraphRunCount,
  })
    ? content.slice(0, -1)
    : content;
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
 * identically-bordered paragraphs. The renderer draws one box around the run:
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

/** Height of a selected acquired line window, including forward clearance
 *  jumps. For each line `li` in `[sliceStart, paintEnd)`:
 *    if (line.topY !== undefined && line.topY > y) y = line.topY;  // float clearance
 *    y += lineHForLine(line);                                       // line box advance
 *  `lineHForLine` is the paragraph-scope ruby/docGrid/line-spacing resolver. */
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

export interface ParaBorderSegment {
  side: 'top' | 'bottom' | 'left' | 'right';
  edge: ParaBorderEdge;
  x1: number;
  y1: number;
  x2: number;
  y2: number;
}

/** ECMA-376 §17.3.1.7 — the vertical extent (in scale-1 points) a paragraph's
 *  BOTTOM border adds BELOW the text box, so following content clears it.
 *
 *  §17.3.1.7 places the bottom border `w:space` points below the text ("the space
 *  after the bottom of the text … before this border is drawn"), and §17.3.4 gives
 *  the border its own width (`w:sz`, eighths of a point). {@link drawParaBorders}
 *  strokes the line centered on `textBottom + space`, so its outer bottom edge is
 *  at `textBottom + space + width/2`. `word-paragraph-border-flow-reservation`
 *  reserves that complete painted extent before the following content.
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
