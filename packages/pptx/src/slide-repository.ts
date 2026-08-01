import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
import {
  BoundedAsyncLruCache,
  type BoundedAsyncLruCacheUsage,
} from '@silurus/ooxml-core/internal/bounded-async-lru-cache';
import { measureStructuralJson } from '@silurus/ooxml-core/internal/resource-measurement';
import { parseResourceLimitError } from '@silurus/ooxml-core/worker';
import type { Slide } from './types.js';

export interface PptxSlideRepositoryOptions {
  /** Immutable slide count for the presentation generation. */
  slideCount: number;
  /** Maximum number of fully parsed slides retained by this repository. */
  maxCachedSlides: number;
  /** Maximum retained structural JSON weight across cached slides. */
  maxCachedStructuralBytes: number;
  /**
   * Produces one complete slide. The producer remains responsible for the hard
   * per-slide limit; this repository governs only resolved-value retention.
   */
  loadSlide: (slideIndex: number) => Slide | PromiseLike<Slide>;
}

/**
 * Presentation-generation-local, bounded ownership for fully parsed slides.
 *
 * Structural JSON bytes are a deterministic retention surrogate, not a claim
 * about engine heap size. The shared cache admits only resolved values, while
 * the format producer bounds the indivisible slide before resolving it. Cached
 * slides are the immutable plain-object snapshots produced by JSON.parse;
 * mutating one after admission would invalidate its recorded structural weight.
 */
export class PptxSlideRepository {
  readonly #slideCount: number;
  readonly #loadSlide: (slideIndex: number) => Slide | PromiseLike<Slide>;
  readonly #cache: BoundedAsyncLruCache<number, Slide>;
  #generation = 0;
  #resourceFailure: OoxmlResourceLimitError | undefined;
  #resourceFailureGeneration: number | undefined;

  constructor(options: PptxSlideRepositoryOptions) {
    if (!Number.isSafeInteger(options.slideCount) || options.slideCount < 0) {
      throw new TypeError('slideCount must be a non-negative safe integer');
    }
    this.#slideCount = options.slideCount;
    this.#loadSlide = options.loadSlide;
    this.#cache = new BoundedAsyncLruCache({
      maxEntries: options.maxCachedSlides,
      maxWeight: options.maxCachedStructuralBytes,
      measure: (slide) => measureStructuralJson(slide).jsonBytes,
    });
  }

  get slideCount(): number {
    return this.#slideCount;
  }

  get usage(): BoundedAsyncLruCacheUsage {
    return this.#cache.usage;
  }

  async load(slideIndex: number): Promise<Slide> {
    this.#assertSlideIndex(slideIndex);
    if (this.#resourceFailure) throw this.#resourceFailure;

    const generation = this.#generation;
    return this.#cache.getOrLoad(slideIndex, async () => {
      let slide: Slide;
      try {
        slide = await this.#loadSlide(slideIndex);
      } catch (error) {
        const resourceFailure = asResourceFailure(error);
        if (!resourceFailure) throw error;
        this.#latchResourceFailure(resourceFailure, generation);
        throw this.#resourceFailureGeneration === generation
          ? (this.#resourceFailure ?? resourceFailure)
          : resourceFailure;
      }

      // A resource failure from another concurrent slide poisons the whole
      // presentation generation. Do not promote this otherwise successful
      // completion after that failure.
      if (this.#resourceFailureGeneration === generation && this.#resourceFailure) {
        throw this.#resourceFailure;
      }
      return slide;
    });
  }

  /**
   * Starts a fresh repository generation. In-flight callers may still receive
   * their value, but a completion from the cleared generation is never cached.
   */
  clear(): void {
    this.#generation += 1;
    this.#resourceFailure = undefined;
    this.#resourceFailureGeneration = undefined;
    this.#cache.clear();
  }

  #assertSlideIndex(slideIndex: number): void {
    if (
      !Number.isSafeInteger(slideIndex) ||
      slideIndex < 0 ||
      slideIndex >= this.#slideCount
    ) {
      throw new RangeError(
        `Slide index ${slideIndex} out of range (count: ${this.#slideCount})`,
      );
    }
  }

  #latchResourceFailure(error: OoxmlResourceLimitError, generation: number): void {
    if (generation !== this.#generation || this.#resourceFailure) return;
    this.#resourceFailure = error;
    this.#resourceFailureGeneration = generation;
    this.#generation += 1;
    // Invalidate every pending token as well as retained slides. A completion
    // racing with the fatal result can therefore never be admitted afterward.
    this.#cache.clear();
  }
}

function asResourceFailure(error: unknown): OoxmlResourceLimitError | undefined {
  return error instanceof OoxmlResourceLimitError
    ? error
    : parseResourceLimitError(error);
}
