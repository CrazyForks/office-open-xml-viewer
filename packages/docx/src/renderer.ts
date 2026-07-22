import type { DocxDocumentModel, BodyElement } from './types';
import type { HyperlinkTarget, LayoutServices, MathRenderer } from './layout/types.js';
import { bodyMathOccurrences, documentMathOccurrences } from './layout/resources.js';
import { paintResourceRegistryOf, privateResourceLookupOf } from './layout/runtime-state.js';
import { selectDocumentLayoutPage } from './layout/document-layout-variants.js';
import { textRunsForPage } from './text-run-projection.js';
import { dropBrowserImageCache } from './paint/browser-images.js';
import { canvasPageScale, renderSelectedDocumentPage } from './paint/canvas-document.js';
import { ensureDocumentLayoutVariants } from './layout/document.js';
import { productionDocumentInput } from './layout/resources.js';
import { prepareBrowserMathResources } from './paint/browser-math.js';
import { createLayoutServices } from './layout-runtime.js';

/** True if any currently representable document story contains OMML. The body
 * array form remains supported for existing callers. */
export function documentHasMath(input: BodyElement[] | DocxDocumentModel): boolean {
  return (Array.isArray(input) ? bodyMathOccurrences(input) : documentMathOccurrences(input)).length > 0;
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
  return prepareBrowserMathResources(productionDocumentInput(input).mathOccurrences, math);
}

/** Information about a rendered text segment for building a transparent selection overlay. */
export interface DocxTextRunInfo {
  /**
   * Authored `w14:paraId` of the source paragraph. Absent when the paragraph
   * does not carry that identifier.
   */
  paragraphId?: string;
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

export function dropColorReplacedCache(
  fetchImage: (path: string, mime: string) => Promise<Blob>,
): void {
  dropBrowserImageCache(fetchImage);
}

function normalizeRenderOptions(
  doc: DocxDocumentModel,
  canvas: HTMLCanvasElement | OffscreenCanvas,
  pageIndex: number,
  options: RenderDocumentOptions,
) {
  const services = options.layoutServices ?? createLayoutServices(
    doc,
    doc.parseError == null ? {
      measureContext: canvas.getContext('2d') as
        | CanvasRenderingContext2D
        | OffscreenCanvasRenderingContext2D
        | null,
    } : {},
  );
  const defaultCurrentDateMs = options.defaultCurrentDateMs ?? Date.now();
  ensureDocumentLayoutVariants(
    services,
    defaultCurrentDateMs,
    () => {
      const productionInput = productionDocumentInput(doc);
      return {
        model: productionInput.document,
        input: productionInput.bodyLayoutInput,
      };
    },
  );
  const selection = selectDocumentLayoutPage(services, {
    currentDate: options.currentDate,
    defaultCurrentDateMs,
  }, pageIndex);
  const scale = canvasPageScale(selection.page, options.width);
  return {
    selection,
    paintOptions: {
      width: options.width,
      dpr: options.dpr,
      defaultTextColor: options.defaultTextColor,
      showTrackChanges: options.showTrackChanges,
      fetchImage: options.fetchImage,
      parseError: doc.parseError != null,
      registry: paintResourceRegistryOf(services),
      privateResources: privateResourceLookupOf<CanvasImageSource>(services),
      textRuns: options.onTextRun
        ? textRunsForPage(selection.layout, pageIndex, { scale })
        : [],
      onTextRun: options.onTextRun,
    },
  };
}

export async function renderDocumentToCanvas(
  doc: DocxDocumentModel,
  canvas: HTMLCanvasElement | OffscreenCanvas,
  pageIndex: number,
  opts: RenderDocumentOptions = {},
): Promise<void> {
  const normalized = normalizeRenderOptions(doc, canvas, pageIndex, opts);
  return renderSelectedDocumentPage(
    normalized.selection.layout,
    normalized.selection.page,
    canvas,
    normalized.paintOptions,
  );
}
