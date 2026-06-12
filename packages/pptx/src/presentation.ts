import type { Presentation, WorkerRequest, WorkerResponse } from './types';
import { renderSlide, type TextRunCallback } from './renderer';
import { createPresentationHandle, type PresentationHandle } from './presentation-handle';
import { selectNotes } from './notes';
import {
  preloadGoogleFonts,
  WorkerBridge,
  defaultDpr,
  type LoadOptions as CoreLoadOptions,
  type MathRenderer,
} from '@silurus/ooxml-core';
import { PPTX_GOOGLE_FONTS } from './google-fonts';
import { findMimeTypeForPath } from './media-mime';
import InlineWorker from './worker.ts?worker&inline';
import wasmAssetUrl from './wasm/pptx_parser_bg.wasm?url';

/** Options for {@link PptxPresentation.load}. The shared load-options type
 *  from `@silurus/ooxml-core` (`useGoogleFonts`, `maxZipEntryBytes`). */
export type LoadOptions = CoreLoadOptions;

/** Options for rendering a single slide onto a canvas. */
export interface RenderSlideOptions {
  /** Display width in CSS pixels. Defaults to canvas.offsetWidth or 960. */
  width?: number;
  /** Device pixel ratio. Defaults to window.devicePixelRatio or 1. */
  dpr?: number;
  /** Called for each rendered text segment. Used to build a transparent text selection overlay. */
  onTextRun?: TextRunCallback;
  /**
   * Skip drawing the play badge overlay on media elements. Used internally by
   * {@link PptxPresentation.presentSlide} so its interactive handle can draw
   * its own play/pause chrome without duplication.
   */
  skipMediaControls?: boolean;
}

/**
 * Headless PPTX rendering engine.
 *
 * Parses `.pptx` archives in a background worker (WASM) but renders slides
 * synchronously on the main thread, so the canvas shares the document's
 * `FontFaceSet` — avoiding subtle wrap differences between system fallback
 * fonts and theme-declared webfonts (e.g. Nunito Sans).
 *
 * Construct via the static `load` factory. A single instance can drive any
 * number of canvases (scroll view, thumbnail grid, master-detail, etc.).
 *
 * @example
 * const pres = await PptxPresentation.load(buffer);
 * await pres.renderSlide(canvas, 0, { width: 960 });
 */
export class PptxPresentation {
  private readonly _worker: Worker;
  private readonly _bridge: WorkerBridge<WorkerResponse>;
  private _presentation: Presentation | null = null;
  private _mediaCache = new Map<string, Promise<Blob>>();
  private _workerReady = false;
  private _workerReadyCallbacks: Array<() => void> = [];
  /** Opt-in OMML equation engine, injected once at {@link load}. Every
   *  `renderSlide` / `presentSlide` reuses it — equations render when present,
   *  and are skipped (engine tree-shaken) when omitted. */
  private _math: MathRenderer | undefined;

  private constructor() {
    this._worker = new InlineWorker();
    this._bridge = new WorkerBridge<WorkerResponse>(this._worker, {
      // The init `ready` handshake carries no id; everything else does.
      correlate: (msg) => ('id' in msg ? msg.id : undefined),
      toError: (msg) => (msg.kind === 'error' ? msg.message : undefined),
      onUnsolicited: (msg) => {
        if (msg.kind === 'ready') {
          this._workerReady = true;
          for (const cb of this._workerReadyCallbacks) cb();
          this._workerReadyCallbacks = [];
        }
      },
    });
    const wasmUrl = new URL(wasmAssetUrl, location.href).href;
    this._bridge.post({ kind: 'init', wasmUrl } satisfies WorkerRequest);
  }

  /** Parse a PPTX from URL or ArrayBuffer. */
  static async load(
    source: string | ArrayBuffer,
    opts: LoadOptions = {},
  ): Promise<PptxPresentation> {
    const pres = new PptxPresentation();
    let buffer: ArrayBuffer;
    if (typeof source === 'string') {
      const res = await fetch(source);
      if (!res.ok) throw new Error(`Failed to fetch: ${res.status} ${res.statusText}`);
      buffer = await res.arrayBuffer();
    } else {
      buffer = source;
    }
    pres._math = opts.math;
    await pres._parse(buffer, opts.maxZipEntryBytes);
    const parsed = pres._presentation;
    if (opts.useGoogleFonts && parsed) {
      // Always load the generic Arabic fallbacks so any Arabic-script run gets
      // a real web font even when its named family is unmapped (the renderer's
      // canvas font stack ends with these two Noto faces).
      await preloadGoogleFonts(
        [parsed.majorFont, parsed.minorFont, 'Noto Naskh Arabic', 'Noto Sans Arabic'],
        PPTX_GOOGLE_FONTS,
      );
    }
    return pres;
  }

  private _waitForWorker(): Promise<void> {
    if (this._workerReady) return Promise.resolve();
    return new Promise((resolve) => this._workerReadyCallbacks.push(resolve));
  }

  private async _parse(buffer: ArrayBuffer, maxZipEntryBytes?: number): Promise<void> {
    await this._waitForWorker();
    const res = await this._bridge.request(
      (id) => ({ kind: 'parse', id, buffer, maxZipEntryBytes }) satisfies WorkerRequest,
      [buffer],
    );
    this._presentation = (res as Extract<WorkerResponse, { kind: 'parsed' }>).presentation;
  }

  /** Total number of slides in the loaded presentation. */
  get slideCount(): number { return this._presentation?.slides.length ?? 0; }

  /** Slide width in EMU. */
  get slideWidth(): number { return this._presentation?.slideWidth ?? 0; }

  /** Slide height in EMU. */
  get slideHeight(): number { return this._presentation?.slideHeight ?? 0; }

  /**
   * Speaker-notes text for a slide (`ppt/notesSlides/notesSlideN.xml`,
   * ECMA-376 §13.3.5 — Notes Slide). Returns the notes-body text as a single
   * string (paragraphs joined with `\n`), or `null` when the slide has no
   * notes part. The notes are parsed at {@link load} time, so this is a
   * synchronous lookup.
   *
   * `slideIndex` is 0-based. Unlike navigation methods it is *not* clamped:
   * an out-of-range or non-integer index returns `null` rather than the notes
   * of the nearest slide (so a tool iterating by index gets an honest "no
   * notes" instead of a duplicated neighbour).
   *
   * @example
   * const pres = await PptxPresentation.load(buffer);
   * for (let i = 0; i < pres.slideCount; i++) {
   *   const notes = pres.getNotes(i);
   *   if (notes) console.log(`Slide ${i + 1} notes:`, notes);
   * }
   */
  getNotes(slideIndex: number): string | null {
    return selectNotes(this._presentation?.slides ?? [], slideIndex);
  }

  /** Render a slide onto the given canvas. */
  async renderSlide(
    canvas: HTMLCanvasElement,
    slideIndex: number,
    opts: RenderSlideOptions = {},
  ): Promise<void> {
    if (!this._presentation) throw new Error('Presentation not loaded');
    const slide = this._presentation.slides[slideIndex];
    if (!slide) throw new Error(`Slide index ${slideIndex} out of range (count: ${this.slideCount})`);
    const dpr = opts.dpr ?? defaultDpr();
    const width = opts.width ?? (canvas.offsetWidth || 960);
    await renderSlide(
      canvas,
      slide,
      this._presentation.slideWidth,
      this._presentation.slideHeight,
      {
        width,
        dpr,
        defaultTextColor: this._presentation.defaultTextColor,
        majorFont: this._presentation.majorFont,
        minorFont: this._presentation.minorFont,
        hlinkColor: this._presentation.hlinkColor ?? null,
        fetchMedia: (path) => this.getMedia(path),
        skipMediaControls: opts.skipMediaControls,
        math: this._math,
      },
      opts.onTextRun,
    );
  }

  /**
   * Extract raw media bytes for a zip path referenced by {@link MediaElement}.
   * Results are cached by path for the lifetime of this instance.
   */
  async getMedia(mediaPath: string): Promise<Blob> {
    const hit = this._mediaCache.get(mediaPath);
    if (hit) return hit;
    const mimeType = this._findMimeTypeForPath(mediaPath);
    const p = (async () => {
      await this._waitForWorker();
      const res = await this._bridge.request(
        (id) => ({ kind: 'extractMedia', id, path: mediaPath }) satisfies WorkerRequest,
      );
      const bytes = (res as Extract<WorkerResponse, { kind: 'mediaExtracted' }>).bytes;
      return new Blob([bytes], { type: mimeType });
    })();
    this._mediaCache.set(mediaPath, p);
    return p;
  }

  private _findMimeTypeForPath(mediaPath: string): string {
    if (!this._presentation) return '';
    return findMimeTypeForPath(this._presentation, mediaPath);
  }

  /**
   * Render a slide and attach canvas-native playback controls for any
   * embedded audio/video. Returns a {@link PresentationHandle} that owns the
   * RAF loop, media elements, and object URLs. Unlike {@link renderSlide}, this
   * method is stateful — always call `handle.destroy()` when leaving the slide.
   */
  async presentSlide(
    canvas: HTMLCanvasElement,
    slideIndex: number,
    opts: RenderSlideOptions = {},
  ): Promise<PresentationHandle> {
    if (!this._presentation) throw new Error('Presentation not loaded');
    const slide = this._presentation.slides[slideIndex];
    if (!slide) throw new Error(`Slide index ${slideIndex} out of range (count: ${this.slideCount})`);
    const dpr = opts.dpr ?? defaultDpr();
    const width = opts.width ?? (canvas.offsetWidth || 960);
    return createPresentationHandle(canvas, slide, {
      width,
      dpr,
      slideWidthEmu: this._presentation.slideWidth,
      fetchMedia: (path) => this.getMedia(path),
      drawBase: () => this.renderSlide(canvas, slideIndex, {
        width,
        dpr,
        skipMediaControls: true,
        onTextRun: opts.onTextRun,
      }),
    });
  }

  /** Terminate the worker and release all resources. */
  destroy(): void {
    this._bridge.terminate();
    this._presentation = null;
    this._mediaCache.clear();
  }
}
