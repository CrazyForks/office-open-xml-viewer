/**
 * Small browser-viewer mechanics shared by the format packages. This module is
 * intentionally an internal subpath: it owns DOM/canvas lifecycle facts, never
 * page/slide/sheet semantics or format state transitions.
 */

export type CanvasRestoreMode = 'display' | 'style-and-bitmap';

export const MAX_NATIVE_TEXT_SELECTION_CHARS = 65_536;
export const MAX_NATIVE_TEXT_SELECTION_LOCATORS = 1_024;

export interface BoundedNativeTextSelection<TLocator> {
  readonly text: string;
  readonly locators: readonly TLocator[];
  readonly truncated: boolean;
  readonly truncationReasons: readonly ('text' | 'runs')[];
  readonly textCharacters: number;
  readonly maxTextCharacters: number;
  readonly maxLocators: number;
}

function boundedSelectionLimit(value: number | undefined, maximum: number, name: string): number {
  const requested = value ?? maximum;
  if (!Number.isFinite(requested) || requested < 0) {
    throw new RangeError(`${name} must be a finite non-negative number.`);
  }
  return Math.min(maximum, Math.floor(requested));
}

function safeUtf16Prefix(value: string, maxCodeUnits: number): string {
  let end = Math.min(value.length, maxCodeUnits);
  if (end > 0 && end < value.length) {
    const previous = value.charCodeAt(end - 1);
    const next = value.charCodeAt(end);
    if (previous >= 0xD800 && previous <= 0xDBFF && next >= 0xDC00 && next <= 0xDFFF) end--;
  }
  return value.slice(0, end);
}

/**
 * Read a browser-native text selection only when every range endpoint belongs
 * to this Viewer. This prevents a cross-DOM selection from leaking adjacent
 * page content into an AI/MCP context. Locators come only from tagged run spans
 * intersected by the native ranges and are detached by the caller's mapper.
 */
export function readBoundedNativeTextSelection<TLocator>(
  root: HTMLElement,
  selection: Selection | null,
  locatorForRun: (run: HTMLElement) => TLocator | null,
  options: Readonly<{ maxChars?: number; maxLocators?: number }> = {},
): BoundedNativeTextSelection<TLocator> | null {
  if (!selection || selection.isCollapsed || selection.rangeCount === 0) return null;
  const selectionSurfaces = [
    ...(root.matches?.('[data-ooxml-selection-surface]') ? [root] : []),
    ...root.querySelectorAll<HTMLElement>('[data-ooxml-selection-surface]'),
  ];
  if (selectionSurfaces.length === 0) return null;
  const isOnSelectionSurface = (node: Node) =>
    selectionSurfaces.some((surface) => surface.contains(node));
  const ranges: Range[] = [];
  for (let index = 0; index < selection.rangeCount; index++) {
    const range = selection.getRangeAt(index);
    if (!root.contains(range.startContainer) || !root.contains(range.endContainer) ||
        !isOnSelectionSurface(range.startContainer) ||
        !isOnSelectionSurface(range.endContainer)) return null;
    ranges.push(range);
  }
  const rawText = selection.toString();
  if (rawText.length === 0) return null;
  const maxChars = boundedSelectionLimit(
    options.maxChars, MAX_NATIVE_TEXT_SELECTION_CHARS, 'maxTextCharacters',
  );
  const maxLocators = boundedSelectionLimit(
    options.maxLocators, MAX_NATIVE_TEXT_SELECTION_LOCATORS, 'maxRunLocators',
  );
  const locators: TLocator[] = [];
  let locatorOverflow = false;
  for (const candidate of root.querySelectorAll<HTMLElement>('[data-ooxml-selection-run]')) {
    if (!ranges.some((range) => {
      try { return range.intersectsNode(candidate); } catch { return false; }
    })) continue;
    const locator = locatorForRun(candidate);
    if (locator === null) continue;
    if (locators.length >= maxLocators) { locatorOverflow = true; break; }
    locators.push(structuredClone(locator));
  }
  if (locators.length === 0 && !locatorOverflow) return null;
  const text = safeUtf16Prefix(rawText, maxChars);
  return {
    text,
    locators,
    truncated: rawText.length > maxChars || locatorOverflow,
    truncationReasons: [
      ...(rawText.length > maxChars ? ['text' as const] : []),
      ...(locatorOverflow ? ['runs' as const] : []),
    ],
    textCharacters: text.length,
    maxTextCharacters: maxChars,
    maxLocators,
  };
}

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

    const ownerDocument = canvas.ownerDocument ?? document;
    this.wrapper = ownerDocument.createElement('div');
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
    const ownerDocument = wrapper.ownerDocument ?? document;
    this.textLayer = enableTextSelection ? ownerDocument.createElement('div') : null;
    if (this.textLayer) {
      this.textLayer.style.cssText = TEXT_LAYER_STYLE;
      wrapper.appendChild(this.textLayer);
    }
    this.highlightLayer = ownerDocument.createElement('div');
    this.highlightLayer.style.cssText = HIGHLIGHT_LAYER_STYLE;
    wrapper.appendChild(this.highlightLayer);
  }
}

export interface BitmapCommitSize {
  readonly cssWidth?: number;
  readonly cssHeight?: number;
}

export interface DestroyableResource {
  destroy(): void;
}

export type CanvasViewerRenderMode = 'main' | 'worker';

/** Resolve the mode of a viewer that may borrow an already-loaded engine. */
export function resolveCanvasViewerMode(
  viewerName: string,
  requestedMode: CanvasViewerRenderMode | undefined,
  engine: Readonly<{ mode: CanvasViewerRenderMode }> | undefined,
): CanvasViewerRenderMode {
  if (engine && requestedMode !== undefined && requestedMode !== engine.mode) {
    throw new Error(
      `${viewerName}: opts.mode='${requestedMode}' conflicts with the borrowed engine's ` +
        `mode='${engine.mode}'. Omit opts.mode when borrowing an engine — ` +
        'the engine owns its render mode.',
    );
  }
  return engine?.mode ?? requestedMode ?? 'main';
}

/**
 * Terminal, generation-safe ownership for a replaceable viewer resource.
 *
 * Page, slide, and sheet viewers share the same lifecycle invariant: a newer
 * acquisition supersedes an older one, a losing candidate is destroyed, and
 * close is permanent. Format-specific rendering remains outside this class.
 */
export class TerminalResourceOwner<T extends DestroyableResource> {
  private generation = 0;
  private resource: T | null;
  private ownsResource: boolean;
  private closed = false;

  constructor(
    private readonly ownerName: string,
    initial: T | null = null,
    ownsInitial = false,
  ) {
    this.resource = initial;
    this.ownsResource = initial !== null && ownsInitial;
  }

  get current(): T | null {
    return this.resource;
  }

  async replace(
    load: () => Promise<T>,
    beforeCommit?: (previous: T | null) => void,
  ): Promise<T | null> {
    this.assertOpen();
    const generation = ++this.generation;
    let candidate: T;
    try {
      candidate = await load();
    } catch (error) {
      if (this.closed) throw this.closedError();
      if (generation !== this.generation) return null;
      throw error;
    }
    if (this.closed) {
      this.dispose(candidate);
      throw this.closedError();
    }
    if (generation !== this.generation) {
      this.dispose(candidate);
      return null;
    }
    try {
      beforeCommit?.(this.resource);
    } catch (error) {
      this.dispose(candidate);
      throw error;
    }
    this.install(candidate, true);
    return candidate;
  }

  install(candidate: T, owned = true): void {
    this.assertOpen();
    // A direct installation is itself a replacement generation. Any loader
    // already in flight must lose when it resolves; otherwise it can overwrite
    // this explicitly installed resource and destroy it as the previous owner.
    this.generation++;
    const previous = this.resource;
    const ownedPrevious = this.ownsResource;
    this.resource = candidate;
    this.ownsResource = owned;
    if (ownedPrevious && previous) this.dispose(previous);
  }

  close(): void {
    if (this.closed) return;
    this.closed = true;
    this.generation++;
    const previous = this.resource;
    const ownedPrevious = this.ownsResource;
    this.resource = null;
    this.ownsResource = false;
    if (ownedPrevious && previous) this.dispose(previous);
  }

  private assertOpen(): void {
    if (this.closed) throw this.closedError();
  }

  private closedError(): Error {
    return new Error(`${this.ownerName} is closed`);
  }

  /** Cleanup cannot change the already-committed ownership transition. */
  private dispose(resource: T): void {
    try { resource.destroy(); } catch {}
  }
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
