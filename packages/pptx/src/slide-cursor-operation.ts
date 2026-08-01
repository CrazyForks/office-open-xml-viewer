import type { OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import {
  decodeOoxmlResourceUsage,
  HARD_MAX_PPTX_SLIDE_JSON_BYTES,
} from '@silurus/ooxml-core/worker';
import type { Slide } from './types.js';

export interface PptxSlideCursorArchive {
  pull_slide(
    slideIndex: number,
    operationId: number,
    generation: number,
    byteCredit: number,
  ): Uint8Array;
  slide_cursor_resource_usage(): Uint8Array;
  acknowledge_slide(operationId: number, generation: number): void;
  cancel_slide(): void | Promise<void>;
  close_presentation_session(): void | Promise<void>;
}

export type PptxSlideAcceptance =
  | void
  | (() => void)
  | { rollback?: () => void; commit?: () => void };

export type PptxSlideArchiveExecutor = <T>(
  operation: (archive: PptxSlideCursorArchive) => T,
) => T;

export interface PptxSlideOperationIdentity {
  readonly operationId: number;
  readonly generation: number;
}

export type PptxSlideAcceptor = (
  slideIndex: number,
  slide: Slide,
  usage?: OoxmlResourceUsageSnapshot,
) => PptxSlideAcceptance;

/** Decode the canonical Rust usage checkpoint without accepting malformed data. */
export function readPptxSlideCursorUsage(
  execute: PptxSlideArchiveExecutor,
): OoxmlResourceUsageSnapshot | undefined {
  try {
    return decodeOoxmlResourceUsage(
      execute((archive) => archive.slide_cursor_resource_usage()),
    );
  } catch (error) {
    // Bootstrap can precede creation of the per-slide ledger. All malformed
    // checkpoints and real parser/resource failures remain fatal.
    if (String(error).includes('slide cursor usage is unavailable')) return undefined;
    throw error;
  }
}

/**
 * Consumer acceptance and Rust ACK form one transaction. The consumer prepares
 * compact/cache state first, Rust commits its cursor journal second, and only
 * then does the consumer publish its prepared state.
 */
export function acknowledgePptxSlide(
  execute: PptxSlideArchiveExecutor,
  identity: PptxSlideOperationIdentity,
  slideIndex: number,
  slide: Slide | undefined,
  acceptSlide?: PptxSlideAcceptor,
): void {
  let rollback: (() => void) | undefined;
  let commit: (() => void) | undefined;
  try {
    if (acceptSlide) {
      if (!slide) throw new Error('slide payload is missing before acknowledgement');
      const accepted = acceptSlide(
        slideIndex,
        slide,
        readPptxSlideCursorUsage(execute),
      );
      if (typeof accepted === 'function') rollback = accepted;
      else if (accepted) ({ rollback, commit } = accepted);
    }
    execute((archive) => archive.acknowledge_slide(
      identity.operationId,
      identity.generation,
    ));
    commit?.();
  } catch (error) {
    try {
      rollback?.();
    } catch {
      // Preserve the parser/acceptance failure that made the transaction fail.
    }
    throw error;
  }
}

/**
 * Worker-local counterpart of the transferred slide pull. It uses the same
 * canonical cursor, hard unit credit, acceptance-before-ACK ordering, and
 * cancellation path while keeping the complete Slide inside the render worker.
 */
export function loadPptxSlideFromCursor(
  execute: PptxSlideArchiveExecutor,
  slideIndex: number,
  identity: PptxSlideOperationIdentity,
  acceptSlide?: PptxSlideAcceptor,
): Slide {
  let slide: Slide | undefined;
  try {
    const bytes = execute((archive) => archive.pull_slide(
      slideIndex,
      identity.operationId,
      identity.generation,
      HARD_MAX_PPTX_SLIDE_JSON_BYTES,
    ));
    slide = JSON.parse(new TextDecoder().decode(bytes)) as Slide;
    acknowledgePptxSlide(execute, identity, slideIndex, slide, acceptSlide);
    return slide;
  } catch (error) {
    try {
      execute((archive) => archive.cancel_slide());
    } catch {
      // Cancellation is best-effort here; keep the original deterministic
      // parser/consumer failure as the public cause.
    }
    slide = undefined;
    throw error;
  }
}
