import { DocxDocument } from './document';
import type { LoadOptions } from './document';
import type { RenderPageOptions } from './types';
import type { DocxTextRunInfo } from './renderer';
import { buildDocxTextLayer } from './text-layer';
import { buildDocxHighlightLayer, type DocxHighlightMatch } from './find-highlight-layer';
import { DocxFindController, type DocxMatchLocation } from './find';
import { openExternalHyperlink } from '@silurus/ooxml-core';
import type { HyperlinkTarget, FindMatch, FindMatchesOptions } from '@silurus/ooxml-core';

export interface DocxViewerOptions extends RenderPageOptions, LoadOptions {
  container?: HTMLElement;
  /**
   * When true, adds a transparent text overlay div over the canvas so the
   * browser's native text selection works on document content.
   */
  enableTextSelection?: boolean;
  /** Called when a page finishes rendering. */
  onPageChange?: (index: number, total: number) => void;
  /** IX1 (design decision — NOT user-confirmed, integrator may veto). Called when
   *  a hyperlink run is clicked. When omitted, the default is: external → open in a
   *  new tab via core `openExternalHyperlink` (sanitised, noopener,noreferrer);
   *  internal → jump to the page whose text contains the bookmark (best-effort). */
  onHyperlinkClick?: (target: HyperlinkTarget) => void;
  /** Called on parse or render errors. */
  onError?: (err: Error) => void;
}

export class DocxViewer {
  private _doc: DocxDocument | null = null;
  private _currentPage = 0;
  private _canvas: HTMLCanvasElement;
  private _wrapper: HTMLDivElement;
  /** The canvas's DOM position BEFORE the constructor reparented it into
   *  {@link _wrapper}, captured so {@link destroy} can return the caller-owned
   *  canvas to exactly where it was. `null` parent = canvas was passed
   *  detached. */
  private _originalParent: Node | null = null;
  private _originalNextSibling: Node | null = null;
  /** The canvas's inline `display` before the constructor forced `block`
   *  (empty string if it was unset), restored on {@link destroy}. */
  private _originalDisplay = '';
  private _textLayer: HTMLDivElement | null = null;
  /** IX2 — the find-highlight overlay layer. Always created (independent of
   *  `enableTextSelection`): highlights ride the same positioned-DOM overlay
   *  mechanism as the selection layer but are visible boxes, not transparent
   *  spans. Sits above the text layer so a highlight shows over a link's hit
   *  region without stealing its clicks (`pointer-events:none`). */
  private _highlightLayer: HTMLDivElement | null = null;
  /** IX2 — find state (per-page runs, matches, active cursor). */
  private _find: DocxFindController;
  /** A 2d context used only to measure text for highlight geometry (its own
   *  1×1 offscreen canvas, so measuring never touches the visible canvas). */
  private _measureCtx: CanvasRenderingContext2D | null = null;
  private _opts: DocxViewerOptions;
  private readonly _mode: 'main' | 'worker';
  /** The canvas's bitmaprenderer context, used only in worker mode (a canvas
   *  holds one context type for its lifetime; the main-mode 2d render path is
   *  never used on the same canvas). */
  private _bitmapCtx: ImageBitmapRenderingContext | null = null;
  private _warnedNoTextSelection = false;
  /** One-shot guard so the worker-mode findText warning prints once. */
  private _warnedNoFind = false;
  /** Set by {@link destroy} (first line). Guards {@link _reportRenderError} so a
   *  render rejection that lands AFTER teardown is swallowed rather than surfaced
   *  to an `onError` / `console.error` on a dead viewer — parity with the scroll
   *  viewers' `_destroyed` flag. */
  private _destroyed = false;
  /**
   * Concurrent-load latch (generation token). Every {@link load} increments this
   * and captures the value; after its engine finishes loading it re-checks the
   * live value and BAILS (destroying its own just-loaded engine) if a newer
   * `load()` has since started. Without it, two overlapping `load(A)`/`load(B)`
   * calls race the WASM parse / worker init, and whichever RESOLVES last wins the
   * swap — even the stale `load(A)` resolving after `load(B)`; the loser's freshly
   * created engine (never installed, or installed then overwritten) then leaks its
   * worker + pinned WASM allocation. The latch composes with SC20: the check runs
   * AFTER the new engine loads but BEFORE the field assignment and
   * `previous?.destroy()`, so a superseded load never touches `this._doc` nor
   * frees the current (newer) engine. {@link destroy} also bumps it so a load in
   * flight at teardown is treated as superseded and its engine cleaned up.
   */
  private _loadGen = 0;

  constructor(canvas: HTMLCanvasElement, opts: DocxViewerOptions = {}) {
    this._canvas = canvas;
    this._opts = opts;
    this._mode = opts.mode ?? 'main';

    // Wrap canvas in a positioned container for the optional text layer overlay
    const parent = canvas.parentElement;
    // Capture the canvas's DOM position and inline display BEFORE reparenting so
    // destroy() can put the caller-owned canvas back exactly where it was.
    this._originalParent = parent;
    this._originalNextSibling = canvas.nextSibling;
    this._originalDisplay = canvas.style.display;
    this._wrapper = document.createElement('div');
    // vertical-align:top removes the inline-block baseline descender gap that
    // otherwise lets the host container's background show through below the
    // canvas (~6 px on default font metrics).
    this._wrapper.style.cssText = 'position:relative;display:inline-block;vertical-align:top;';
    // Force `display:block` on the canvas so it does not inherit the inline
    // baseline of the wrapper, which would otherwise leave a 4–6px descender
    // gap between the canvas bottom and the wrapper bottom — the host
    // container's background would show through that strip.
    if (!canvas.style.display) canvas.style.display = 'block';
    if (parent) {
      parent.insertBefore(this._wrapper, canvas);
    }
    this._wrapper.appendChild(canvas);

    // Worker mode paints worker-produced bitmaps via a bitmaprenderer context,
    // grabbed once (a canvas holds one context type for its lifetime).
    if (this._mode === 'worker') {
      this._bitmapCtx = canvas.getContext('bitmaprenderer');
    }

    if (opts.enableTextSelection) {
      this._textLayer = document.createElement('div');
      this._textLayer.style.cssText =
        'position:absolute;top:0;left:0;width:100%;height:100%;' +
        'overflow:hidden;pointer-events:none;user-select:text;-webkit-user-select:text;';
      this._wrapper.appendChild(this._textLayer);
    }

    // IX2 — the find-highlight overlay layer. Appended last so it stacks above
    // the text/selection layer; `pointer-events:none` keeps selection + link
    // clicks working through it. In worker mode `onTextRun` can't fire, so it
    // stays empty (find is main-mode only, warned at find() time).
    this._highlightLayer = document.createElement('div');
    this._highlightLayer.style.cssText =
      'position:absolute;top:0;left:0;width:100%;height:100%;' +
      'overflow:hidden;pointer-events:none;';
    this._wrapper.appendChild(this._highlightLayer);

    this._find = new DocxFindController(
      () => this.pageCount,
      (page) => this._collectPageRuns(page),
    );
  }

  /**
   * Load a DOCX from URL or ArrayBuffer and render the first page.
   *
   * Error contract (shared by all three viewers):
   * - Parse/load failure (the underlying `DocxDocument.load()` call itself
   *   rejects): if an `onError` callback was provided it is invoked and `load`
   *   resolves normally; if not, the error is rethrown so it is never silently
   *   swallowed.
   * - Render failure (the first page fails to draw AFTER a successful
   *   parse/load): routed to the shared `_reportRenderError` contract (`onError`
   *   if provided, else `console.error` — never silent) and `load` still
   *   RESOLVES, matching every subsequent navigation call.
   */
  async load(source: string | ArrayBuffer): Promise<void> {
    // SC20 atomic swap: retain the previous engine locally and only tear it down
    // AFTER the new one loads successfully. A re-load thus never orphans the old
    // engine's worker + pinned WASM allocation (the leak this guards), yet a
    // FAILED re-load keeps the current document + its rendered page intact rather
    // than dropping to an empty viewer. The 2× memory window is bounded to the
    // load itself (the old engine is freed the moment the new model arrives).
    const gen = ++this._loadGen;
    const previous = this._doc;
    try {
      const doc = await DocxDocument.load(source, {
        useGoogleFonts: this._opts.useGoogleFonts,
        maxZipEntryBytes: this._opts.maxZipEntryBytes,
        workerTimeoutMs: this._opts.workerTimeoutMs,
        wasmUrl: this._opts.wasmUrl,
        math: this._opts.math,
        mode: this._mode,
      });
      if (gen !== this._loadGen) {
        // A newer load() (or destroy()) started while this one was in flight — we
        // lost the concurrent-load race. Destroy the engine we just loaded (it was
        // never installed) and leave the winning load's engine + SC20 swap
        // untouched: do NOT touch `this._doc` and do NOT destroy `previous`
        // (irrelevant to the winner; possibly already stale).
        doc.destroy();
        return;
      }
      this._doc = doc;
      previous?.destroy();
      this._currentPage = 0;
      // A new document invalidates any prior find state (cached runs / matches).
      this._find.invalidate();
      await this._render();
    } catch (err) {
      // Superseded loads own no error reporting — the winning load (or destroy())
      // is the outcome the caller awaits; swallow this stale rejection.
      if (gen !== this._loadGen) return;
      const e = err instanceof Error ? err : new Error(String(err));
      if (this._opts.onError) {
        this._opts.onError(e);
        return;
      }
      throw e;
    }
  }

  get pageCount(): number {
    return this._doc?.pageCount ?? 0;
  }

  get currentPage(): number {
    return this._currentPage;
  }

  /** The underlying <canvas> element. */
  get canvasElement(): HTMLCanvasElement {
    return this._canvas;
  }

  async goToPage(index: number): Promise<void> {
    if (!this._doc) return;
    const clamped = Math.max(0, Math.min(index, this.pageCount - 1));
    this._currentPage = clamped;
    await this._render();
  }

  async nextPage(): Promise<void> { await this.goToPage(this._currentPage + 1); }
  async prevPage(): Promise<void> { await this.goToPage(this._currentPage - 1); }

  /**
   * IX2 — find every occurrence of `query` in the document and highlight them
   * all (a soft box per match, drawn on the highlight overlay over the drawn
   * glyphs). Returns every match in document order, each tagged with its
   * `{ page }` (0-based). Case-insensitive by default (browser find-in-page);
   * pass `{ caseSensitive: true }` to match case exactly.
   *
   * Scans all pages, so a large document renders each page once offscreen to
   * read its text (the visible page reuses its on-screen render). Not available
   * in `mode: 'worker'` (the `onTextRun` stream can't cross the worker boundary)
   * — returns `[]` there after a one-time warning. An empty query clears the
   * find and returns `[]`.
   */
  async findText(
    query: string,
    opts: FindMatchesOptions = {},
  ): Promise<FindMatch<DocxMatchLocation>[]> {
    if (!this._doc) return [];
    if (this._mode === 'worker') {
      if (!this._warnedNoFind) {
        this._warnedNoFind = true;
        console.warn(
          "[ooxml] findText is unavailable in mode: 'worker' (text runs can't cross the worker boundary). Use mode: 'main'.",
        );
      }
      return [];
    }
    const matches = await this._find.find(query, opts);
    // Redraw the current page's highlights (matches on it become visible without
    // navigating). Cheap DOM geometry — no page re-render.
    this._redrawHighlights();
    return matches;
  }

  /**
   * IX2 — move to the next match (wrap-around from last to first), navigating to
   * its page if needed, and draw it in the distinct active-match colour. Returns
   * the now-active match, or `null` when there are no matches. Call
   * {@link findText} first.
   */
  async findNext(): Promise<FindMatch<DocxMatchLocation> | null> {
    return this._activateMatch(this._find.next());
  }

  /** IX2 — move to the previous match (wrap-around from first to last). */
  async findPrev(): Promise<FindMatch<DocxMatchLocation> | null> {
    return this._activateMatch(this._find.prev());
  }

  /** IX2 — clear all highlights and reset the find state. */
  clearFind(): void {
    this._find.invalidate();
    this._redrawHighlights();
  }

  /** Navigate to the active match's page (if not already there) and redraw the
   *  highlights so the active box shows in the emphasis colour. */
  private async _activateMatch(
    match: FindMatch<DocxMatchLocation> | null,
  ): Promise<FindMatch<DocxMatchLocation> | null> {
    if (!match) {
      this._redrawHighlights();
      return null;
    }
    if (match.location.page !== this._currentPage) {
      // goToPage re-renders, which rebuilds the highlight layer for the new page.
      await this.goToPage(match.location.page);
    } else {
      this._redrawHighlights();
    }
    return match;
  }

  /** Rebuild the highlight overlay for the current page from cached runs
   *  (no page re-render). */
  private _redrawHighlights(): void {
    const runs = this._find.pageRuns(this._currentPage) ?? [];
    this._buildHighlightLayer(runs);
  }

  /**
   * Terminate the parser worker and release resources.
   *
   * The caller-owned `<canvas>` is returned to the DOM position it held before
   * the constructor was called (same parent, same next-sibling) and its inline
   * `display` is restored, so the canvas can be reused — e.g. to construct a new
   * viewer on the same element. If the canvas was passed detached (no parent) it
   * is simply removed from the internal wrapper. Safe to call more than once.
   */
  destroy(): void {
    // First line: block any render rejection racing in from surfacing on a dead
    // viewer (checked at the top of _reportRenderError). Bump the load generation
    // too so a load() still in flight is treated as superseded and its engine is
    // cleaned up rather than installed onto a torn-down viewer.
    this._destroyed = true;
    this._loadGen++;
    this._doc?.destroy();
    this._doc = null;
    // IX2 — drop the find state (matches + cached runs) so a stale
    // findNext()/findPrev() after teardown returns null instead of a match
    // pointing into a dead viewer.
    this._find.invalidate();
    // Return the caller-owned canvas to its original DOM slot before discarding
    // the wrapper. insertBefore still works if the original parent was itself
    // detached; when there was no original parent the canvas is left detached
    // (just pulled out of the wrapper). The recorded next-sibling may have been
    // removed or moved by the caller since construction — insertBefore throws
    // NotFoundError for a reference that is no longer a child of the parent, so
    // fall back to appending at the end in that case.
    if (this._originalParent) {
      const ref =
        this._originalNextSibling && this._originalNextSibling.parentNode === this._originalParent
          ? this._originalNextSibling
          : null;
      this._originalParent.insertBefore(this._canvas, ref);
    } else if (this._canvas.parentNode) {
      this._canvas.parentNode.removeChild(this._canvas);
    }
    this._canvas.style.display = this._originalDisplay;
    this._wrapper.remove();
  }

  private async _render(): Promise<void> {
    // PD14 render-error contract (shared with the scroll viewers): navigation
    // (`nextPage`/`prevPage`/`goToPage`) is often called `void`-style, so an
    // unguarded throw would surface as an unhandled promise rejection. Catch here
    // and route to `onError` (or `console.error` — never silent) so a page render
    // failure is handled the same way in `load()` and every navigation.
    try {
      await this._renderPage();
    } catch (err) {
      this._reportRenderError(err);
    }
  }

  /** Route a render failure to `onError`, or `console.error` when none is given
   *  (never fully silent), and never after teardown. Mirrors the scroll viewers'
   *  `_reportRenderError`. */
  private _reportRenderError(err: unknown): void {
    if (this._destroyed) return;
    const e = err instanceof Error ? err : new Error(String(err));
    if (this._opts.onError) this._opts.onError(e);
    else console.error('[ooxml] DocxViewer render failed:', e);
  }

  private async _renderPage(): Promise<void> {
    if (!this._doc) return;
    const isWorker = this._mode === 'worker';
    // In worker mode rendering happens off the main thread, so the onTextRun
    // callback can't fire — the text-selection overlay is unavailable.
    if (isWorker && this._textLayer && !this._warnedNoTextSelection) {
      this._warnedNoTextSelection = true;
      console.warn(
        "[ooxml] text selection is unavailable in mode: 'worker'; the overlay will be empty. Use mode: 'main' for selectable text.",
      );
    }
    if (isWorker) {
      const dpr = this._opts.dpr ?? (typeof window !== 'undefined' ? window.devicePixelRatio || 1 : 1);
      // Only serializable render options may cross to the worker — spreading the
      // full viewer opts would postMessage non-cloneable values (the math
      // engine, callbacks, container element) and throw a DataCloneError.
      const bmp = await this._doc.renderPageToBitmap(this._currentPage, {
        width: this._opts.width,
        dpr: this._opts.dpr,
        defaultTextColor: this._opts.defaultTextColor,
        showTrackChanges: this._opts.showTrackChanges,
      });
      this._canvas.width = bmp.width;
      this._canvas.height = bmp.height;
      // The bitmap is sized in device px; mirror the main renderer by setting
      // the CSS size to the logical (÷dpr) dimensions so it isn't 2× on HiDPI.
      this._canvas.style.width = `${Math.round(bmp.width / dpr)}px`;
      this._canvas.style.height = `${Math.round(bmp.height / dpr)}px`;
      this._bitmapCtx?.transferFromImageBitmap(bmp);
    } else {
      // Collect runs unconditionally (not just when a text layer exists): the
      // find-highlight overlay needs the current page's run geometry too, and
      // caching them here means find() reuses the visible render for this page
      // instead of re-rendering it offscreen.
      const runs: DocxTextRunInfo[] = [];
      const onTextRun = (r: DocxTextRunInfo) => runs.push(r);
      await this._doc.renderPage(this._canvas, this._currentPage, { ...this._opts, onTextRun });
      if (this._textLayer) {
        this._buildTextLayer(this._textLayer, runs);
      }
      // Feed the just-rendered page's runs to the find controller so highlight
      // geometry matches exactly what was drawn, then (re)draw the highlights.
      this._find.setPageRuns(this._currentPage, runs);
      this._buildHighlightLayer(runs);
    }
    this._opts.onPageChange?.(this._currentPage, this.pageCount);
  }

  /** Draw the find-highlight boxes for the current page from its runs. Clears
   *  the overlay when there is no active find. */
  private _buildHighlightLayer(runs: DocxTextRunInfo[]): void {
    const layer = this._highlightLayer;
    if (!layer) return;
    const cssWidth = this._canvas.style.width || this._canvas.width + 'px';
    const cssHeight = this._canvas.style.height || this._canvas.height + 'px';
    const highlights: DocxHighlightMatch[] = this._find.pageHighlights(this._currentPage);
    buildDocxHighlightLayer(
      layer,
      runs,
      highlights,
      cssWidth,
      cssHeight,
      (font) => this._measureForFont(font),
    );
  }

  /** A width-measurer primed with `font`, backed by a private 1×1 canvas so it
   *  never disturbs the visible canvas's context state. */
  private _measureForFont(font: string): (s: string) => number {
    if (!this._measureCtx) {
      const c = document.createElement('canvas');
      this._measureCtx = c.getContext('2d');
    }
    const ctx = this._measureCtx;
    if (!ctx) return (s) => s.length; // measurement unavailable (headless w/o canvas)
    ctx.font = font;
    return (s) => ctx.measureText(s).width;
  }

  /** Render a page to a throwaway offscreen canvas purely to collect its runs
   *  (text + geometry) for search, without touching the visible canvas. Used by
   *  the find controller for pages other than the one on screen. */
  private async _collectPageRuns(page: number): Promise<DocxTextRunInfo[]> {
    if (!this._doc || this._mode === 'worker') return [];
    // The currently displayed page's runs are already cached by _renderPage; the
    // controller only calls this for other pages.
    const runs: DocxTextRunInfo[] = [];
    const off =
      typeof OffscreenCanvas !== 'undefined'
        ? new OffscreenCanvas(1, 1)
        : (document.createElement('canvas') as HTMLCanvasElement);
    await this._doc.renderPage(off, page, {
      ...this._opts,
      onTextRun: (r: DocxTextRunInfo) => runs.push(r),
    });
    return runs;
  }

  private _buildTextLayer(layer: HTMLDivElement, runs: DocxTextRunInfo[]): void {
    buildDocxTextLayer(
      layer,
      runs,
      this._canvas.style.width || this._canvas.width + 'px',
      this._canvas.style.height || this._canvas.height + 'px',
      this._hyperlinkHandler(),
    );
  }

  /**
   * IX1/IX-nav — the click handler passed to the text-layer overlay. When the
   * caller supplied `onHyperlinkClick`, it fully owns the behaviour (the default
   * is suppressed). Otherwise the built-in default is: an external link opens in
   * a new tab through core `openExternalHyperlink` (URL sanitised against the
   * safe scheme allowlist, `noopener,noreferrer`); an internal `<w:anchor>` link
   * resolves its bookmark name to a page via
   * {@link DocxDocument.getBookmarkPage} (ECMA-376 §17.16.23) and jumps there
   * with {@link goToPage}. An anchor naming no known bookmark is a safe no-op
   * rather than a jump to a guessed page.
   */
  private _hyperlinkHandler(): (target: HyperlinkTarget) => void {
    const custom = this._opts.onHyperlinkClick;
    if (custom) return custom;
    return (target: HyperlinkTarget): void => {
      if (target.kind === 'external') {
        openExternalHyperlink(target.url);
        return;
      }
      // Internal anchor (IX-nav): map the bookmark name to its destination page
      // and navigate. `undefined` ⇒ no bookmark of that name ⇒ inert.
      const page = this._doc?.getBookmarkPage(target.ref);
      if (page !== undefined) void this.goToPage(page);
    };
  }
}
