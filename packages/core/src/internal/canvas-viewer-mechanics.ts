/**
 * Small browser-viewer mechanics shared by the format packages. This module is
 * intentionally an internal subpath: it owns DOM/canvas lifecycle facts, never
 * page/slide/sheet semantics or format state transitions.
 */

export type CanvasRestoreMode = 'display' | 'style-and-bitmap';

export interface CallerCanvasMountOptions {
  readonly wrapperCssText: string;
  readonly forceDisplayBlock?: boolean;
  readonly restoreMode?: CanvasRestoreMode;
}

/** Reparents a caller-owned canvas into one wrapper and restores it exactly. */
export class CallerCanvasMount {
  readonly wrapper: HTMLDivElement;

  private readonly originalParent: Node | null;
  private readonly originalNextSibling: Node | null;
  private readonly originalDisplay: string;
  private readonly originalStyle: string | null;
  private readonly originalWidth: number;
  private readonly originalHeight: number;
  private restored = false;

  constructor(
    readonly canvas: HTMLCanvasElement,
    private readonly options: CallerCanvasMountOptions,
  ) {
    this.originalParent = canvas.parentNode;
    this.originalNextSibling = canvas.nextSibling;
    this.originalDisplay = canvas.style.display;
    this.originalStyle = options.restoreMode === 'style-and-bitmap'
      ? canvas.getAttribute('style')
      : null;
    this.originalWidth = canvas.width;
    this.originalHeight = canvas.height;

    this.wrapper = document.createElement('div');
    this.wrapper.style.cssText = options.wrapperCssText;
    if (options.forceDisplayBlock && !canvas.style.display) canvas.style.display = 'block';
    if (this.originalParent) this.originalParent.insertBefore(this.wrapper, canvas);
    this.wrapper.appendChild(canvas);
  }

  /** Idempotently restore the original DOM slot and configured canvas state. */
  restore(): void {
    if (this.restored) return;
    this.restored = true;

    if (this.originalParent) {
      const reference = this.originalNextSibling?.parentNode === this.originalParent
        ? this.originalNextSibling
        : null;
      this.originalParent.insertBefore(this.canvas, reference);
    } else if (this.canvas.parentNode) {
      this.canvas.parentNode.removeChild(this.canvas);
    }

    if ((this.options.restoreMode ?? 'display') === 'style-and-bitmap') {
      if (this.originalStyle === null) this.canvas.removeAttribute('style');
      else this.canvas.setAttribute('style', this.originalStyle);
      this.canvas.width = this.originalWidth;
      this.canvas.height = this.originalHeight;
    } else {
      this.canvas.style.display = this.originalDisplay;
    }
    this.wrapper.remove();
  }
}

const TEXT_LAYER_STYLE =
  'position:absolute;top:0;left:0;width:100%;height:100%;' +
  'overflow:hidden;pointer-events:none;user-select:text;-webkit-user-select:text;';
const HIGHLIGHT_LAYER_STYLE =
  'position:absolute;top:0;left:0;width:100%;height:100%;' +
  'overflow:hidden;pointer-events:none;';

/** DOM containers only; each format remains responsible for overlay contents. */
export class CanvasOverlayHost {
  readonly textLayer: HTMLDivElement | null;
  readonly highlightLayer: HTMLDivElement;

  constructor(wrapper: HTMLElement, enableTextSelection: boolean) {
    this.textLayer = enableTextSelection ? document.createElement('div') : null;
    if (this.textLayer) {
      this.textLayer.style.cssText = TEXT_LAYER_STYLE;
      wrapper.appendChild(this.textLayer);
    }
    this.highlightLayer = document.createElement('div');
    this.highlightLayer.style.cssText = HIGHLIGHT_LAYER_STYLE;
    wrapper.appendChild(this.highlightLayer);
  }
}

export interface BitmapCommitSize {
  readonly cssWidth?: number;
  readonly cssHeight?: number;
}

/**
 * Generation gate for a single static canvas. Format renderers still perform
 * their own drawing; this object prevents stale completion side effects and
 * owns worker ImageBitmap commit/disposal.
 */
export class StaticCanvasRenderDispatcher {
  private generation = 0;
  private destroyed = false;
  private readonly bitmapContext: ImageBitmapRenderingContext | null;

  constructor(
    private readonly canvas: HTMLCanvasElement,
    bitmapMode: boolean,
  ) {
    this.bitmapContext = bitmapMode ? canvas.getContext('bitmaprenderer') : null;
  }

  begin(): number {
    return ++this.generation;
  }

  isCurrent(generation: number): boolean {
    return !this.destroyed && generation === this.generation;
  }

  /** Commit a worker bitmap only if it still belongs to the active generation. */
  commitBitmap(
    generation: number,
    bitmap: ImageBitmap,
    size: BitmapCommitSize = {},
  ): boolean {
    if (!this.isCurrent(generation)) {
      bitmap.close();
      return false;
    }
    if (!this.bitmapContext) {
      bitmap.close();
      throw new Error('bitmaprenderer context not available');
    }
    if (this.canvas.width !== bitmap.width) this.canvas.width = bitmap.width;
    if (this.canvas.height !== bitmap.height) this.canvas.height = bitmap.height;
    if (size.cssWidth !== undefined) this.canvas.style.width = `${size.cssWidth}px`;
    if (size.cssHeight !== undefined) this.canvas.style.height = `${size.cssHeight}px`;
    try {
      this.bitmapContext.transferFromImageBitmap(bitmap);
    } catch (error) {
      bitmap.close();
      throw error;
    }
    return true;
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    this.generation++;
  }
}

/** Shared render-error delivery, with a permanent close gate for teardown. */
export class CanvasViewerErrorRouter {
  private closed = false;

  constructor(
    private readonly viewerName: string,
    private readonly onError?: (error: Error) => void,
  ) {}

  report(error: unknown): void {
    if (this.closed) return;
    const normalized = error instanceof Error ? error : new Error(String(error));
    if (this.onError) this.onError(normalized);
    else console.error(`[ooxml] ${this.viewerName} render failed:`, normalized);
  }

  close(): void {
    this.closed = true;
  }
}
