import {
  acquireBitmapCacheLease,
  clampCanvasSize,
  defaultDpr,
  isHTMLCanvas,
  PT_TO_PX,
} from '@silurus/ooxml-core';
import type { Duotone } from '@silurus/ooxml-core';
import type {
  DocumentLayout,
  LayoutPage,
  PaintResourceRegistry,
} from '../layout/types.js';
import { preloadPaintImages, imageKey, type DocxFetchImage } from './browser-images.js';
import {
  createCanvasPaintResourcePainter,
  paintLayoutPageContent,
  paintLayoutPage,
} from './canvas-page.js';
import { canonicalCanvasPaintResourceHandlers } from './canonical-resource-handlers.js';
import {
  createProductionPaintResourceSession,
  unavailablePaintResourceHandle,
} from './resource-session.js';
import type { PaintCanvas2D } from './types.js';

interface PrivatePaintResourceLookup {
  readonly keys: readonly string[];
  resolve(resourceKey: string): CanvasImageSource;
}

export interface CanvasDocumentPaintOptions<TTextRun> {
  readonly width?: number;
  readonly dpr?: number;
  readonly defaultTextColor?: string;
  readonly showTrackChanges?: boolean;
  readonly fetchImage?: DocxFetchImage;
  readonly parseError: boolean;
  readonly registry: PaintResourceRegistry;
  readonly privateResources?: PrivatePaintResourceLookup;
  readonly textRuns: readonly TTextRun[];
  readonly onTextRun?: (run: TTextRun) => void;
}

/** Per-canvas cancellation token: only the newest asynchronous image preload
 * may paint after rapid navigation reuses the same canvas. */
const renderTokens = new WeakMap<HTMLCanvasElement | OffscreenCanvas, number>();

export function canvasPageScale(page: LayoutPage, width?: number): number {
  return (width ?? page.geometry.widthPt * PT_TO_PX) / page.geometry.widthPt;
}

export async function renderSelectedDocumentPage<TTextRun>(
  layout: DocumentLayout,
  page: LayoutPage,
  canvas: HTMLCanvasElement | OffscreenCanvas,
  options: CanvasDocumentPaintOptions<TTextRun>,
): Promise<void> {
  const releaseLease = options.fetchImage
    ? acquireBitmapCacheLease(options.fetchImage)
    : undefined;
  try {
    const token = (renderTokens.get(canvas) ?? 0) + 1;
    renderTokens.set(canvas, token);
    const superseded = (): boolean => renderTokens.get(canvas) !== token;
    const dpr = options.dpr ?? defaultDpr();
    const paintCanvas: HTMLCanvasElement | OffscreenCanvas =
      page.layers.capabilities.requiresElementBackedVerticalGlyphPaint
      && !isHTMLCanvas(canvas)
      && typeof document !== 'undefined'
        ? document.createElement('canvas')
        : canvas;
    const context = paintCanvas.getContext('2d') as PaintCanvas2D | null;
    if (!context) throw new Error('2D canvas is unavailable for DOCX paint');
    const scale = canvasPageScale(page, options.width);
    const cssWidth = page.geometry.widthPt * scale;
    const cssHeight = page.geometry.heightPt * scale;
    const clamped = clampCanvasSize(cssWidth * dpr, cssHeight * dpr);
    const effectiveDpr = clamped.clamped ? dpr * clamped.scale : dpr;
    canvas.width = clamped.width;
    canvas.height = clamped.height;
    if (paintCanvas !== canvas) {
      paintCanvas.width = clamped.width;
      paintCanvas.height = clamped.height;
    }
    if (isHTMLCanvas(canvas)) {
      canvas.style.width = `${cssWidth}px`;
      canvas.style.height = `${cssHeight}px`;
      if (!canvas.style.display) canvas.style.display = 'block';
    }
    if (isHTMLCanvas(paintCanvas) && paintCanvas !== canvas) {
      paintCanvas.style.width = `${cssWidth}px`;
      paintCanvas.style.height = `${cssHeight}px`;
    }
    context.scale(effectiveDpr, effectiveDpr);
    context.fillStyle = '#ffffff';
    context.fillRect(0, 0, cssWidth, cssHeight);

    if (options.parseError) {
      await paintLayoutPage(layout, 0, canvas, { scale, dpr: effectiveDpr });
      return;
    }

    let images;
    try {
      images = await preloadPaintImages(options.registry.descriptors, options.fetchImage);
    } catch (error) {
      if (superseded()) return;
      throw error;
    }
    if (superseded()) return;

    const session = createProductionPaintResourceSession(options.registry, (descriptor) => {
      if (descriptor.kind === 'math') {
        return options.privateResources?.keys.includes(descriptor.resourceKey)
          ? options.privateResources.resolve(descriptor.resourceKey)
          : unavailablePaintResourceHandle('optional math renderer unavailable');
      }
      if (descriptor.kind === 'image' || descriptor.kind === 'picture-bullet') {
        return images.get(imageKey(
          descriptor.partPath,
          descriptor.colorReplaceFrom,
          descriptor.duotone as Duotone | undefined,
        )) ?? unavailablePaintResourceHandle(
          options.fetchImage
            ? 'unsupported image format produced no drawable output'
            : 'image byte source unavailable',
        );
      }
      return undefined;
    });
    const resources = createCanvasPaintResourcePainter(
      session,
      canonicalCanvasPaintResourceHandlers,
    );
    context.save();
    try {
      context.scale(scale, scale);
      paintLayoutPageContent(page, {
        ctx: context,
        scale,
        dpr: effectiveDpr,
        resources,
        defaultTextColor: options.defaultTextColor ?? '#000000',
        showTrackChanges: options.showTrackChanges ?? true,
      });
    } finally {
      context.restore();
    }
    if (paintCanvas !== canvas) {
      if (superseded()) return;
      const destination = canvas.getContext('2d') as PaintCanvas2D | null;
      if (!destination) throw new Error('2D canvas is unavailable for DOCX paint projection');
      destination.drawImage(paintCanvas, 0, 0);
    }
    if (options.onTextRun) {
      for (const run of options.textRuns) options.onTextRun(run);
    }
  } finally {
    releaseLease?.();
  }
}
