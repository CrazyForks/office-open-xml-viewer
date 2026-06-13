import type { MediaElement, WorkerResponse } from './types';

/** Lightweight summary returned by the render worker's `parse` — everything
 *  the main-thread proxy needs for its synchronous getters. The full model
 *  stays in the worker. */
export interface PresentationMeta {
  slideCount: number;
  slideWidth: number;
  slideHeight: number;
  majorFont: string | null;
  minorFont: string | null;
  /** Speaker-notes text per slide (same contract as PptxPresentation.getNotes). */
  notes: (string | null)[];
  /** Media elements per slide (geometry + paths), for main-thread playback
   *  overlays. Small: a handful of plain objects per slide. */
  mediaElements: MediaElement[][];
}

// The base `parse` arm from types.ts is intentionally NOT reused: the render
// worker's `parse` carries an extra `useGoogleFonts` flag, and two `parse`
// arms in one union would defeat `kind`-based narrowing at use sites. The
// `init` / `extractMedia` arms are copied verbatim from `WorkerRequest`.
export type RenderWorkerRequest =
  | { kind: 'init'; wasmUrl: string }
  | { kind: 'extractMedia'; id: number; path: string }
  | { kind: 'parse'; id: number; buffer: ArrayBuffer; maxZipEntryBytes?: number; useGoogleFonts?: boolean }
  | { kind: 'renderSlide'; id: number; slideIndex: number; width: number; dpr: number; skipMediaControls?: boolean };

export type RenderWorkerResponse =
  | Exclude<WorkerResponse, { kind: 'parsed' }>
  | { kind: 'parsedMeta'; id: number; meta: PresentationMeta }
  | { kind: 'slideRendered'; id: number; bitmap: ImageBitmap };
