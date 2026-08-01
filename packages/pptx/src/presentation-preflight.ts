import {
  OoxmlResourceLimitError,
  type OoxmlResourceUsageSnapshot,
} from '@silurus/ooxml-core';
import {
  cappedAdd,
  measureStructuralJson,
} from '@silurus/ooxml-core/internal/resource-measurement';
import { HARD_MAX_PPTX_PREFLIGHT_PROJECTION_BYTES } from '@silurus/ooxml-core/worker';
import { PptxFontPreloadAccumulator } from './google-fonts';
import type { MediaElement, Slide } from './types';
import type {
  PresentationBootstrap,
  PresentationBootstrapSlide,
} from './worker-protocol';

const ZERO_RESOURCE_USAGE: OoxmlResourceUsageSnapshot = Object.freeze({
  archiveEntryCount: 0,
  declaredInflatedBytes: 0,
  distinctInflatedBytes: 0,
  operationInflatedBytes: 0,
});

export const PPTX_MAX_PREFLIGHT_PROJECTION_BYTES =
  HARD_MAX_PPTX_PREFLIGHT_PROJECTION_BYTES;

export interface PresentationPreflightSlide {
  readonly index: number;
  readonly partName?: string;
  readonly notes: string | null;
  readonly hidden: boolean;
  readonly mediaElements: readonly Readonly<MediaElement>[];
}

/**
 * Compact immutable facts retained for synchronous viewer behavior while full
 * slides are pulled and cached independently. Every field is a direct
 * projection of the canonical Rust Slide/PresentationML model: notes
 * (ECMA-376 Part 1 §13.3.5), hidden state (`p:sld@show`, §19.3.1.38),
 * media geometry/relationships, and `sldIdLst` OPC part identity.
 */
export interface PresentationPreflight {
  readonly slideCount: number;
  readonly slideWidth: number;
  readonly slideHeight: number;
  readonly defaultTextColor: string | null;
  readonly majorFont: string | null;
  readonly minorFont: string | null;
  readonly hlinkColor: string | null;
  readonly folHlinkColor: string | null;
  readonly slides: readonly PresentationPreflightSlide[];
  readonly fontPreloadNames: readonly (string | null)[];
}

function assertNullableString(value: unknown, field: string): asserts value is string | null {
  if (value !== null && typeof value !== 'string') {
    throw new Error(`invalid PPTX presentation bootstrap ${field}`);
  }
}

function copyBootstrapSlide(
  value: unknown,
  expectedIndex: number,
): PresentationBootstrapSlide {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`invalid PPTX presentation bootstrap slide at ${expectedIndex}`);
  }
  const candidate = value as Partial<PresentationBootstrapSlide>;
  if (candidate.index !== expectedIndex) {
    throw new Error(`invalid PPTX presentation bootstrap slide index ${candidate.index}`);
  }
  if (candidate.partName !== undefined && typeof candidate.partName !== 'string') {
    throw new Error(`invalid PPTX presentation bootstrap slide partName at ${expectedIndex}`);
  }
  return Object.freeze({
    index: candidate.index,
    ...(candidate.partName === undefined ? {} : { partName: candidate.partName }),
  });
}

/** Validate and detach the JSON-decoded Rust bootstrap from mutable callers. */
export function normalizePresentationBootstrap(
  value: unknown,
): PresentationBootstrap {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error('invalid PPTX presentation bootstrap payload');
  }
  const candidate = value as Partial<PresentationBootstrap>;
  if (
    !Number.isSafeInteger(candidate.slideCount) || (candidate.slideCount ?? -1) < 0 ||
    !Number.isSafeInteger(candidate.slideWidth) || (candidate.slideWidth ?? 0) <= 0 ||
    !Number.isSafeInteger(candidate.slideHeight) || (candidate.slideHeight ?? 0) <= 0 ||
    !Array.isArray(candidate.slides) ||
    candidate.slides.length !== candidate.slideCount
  ) {
    throw new Error('invalid PPTX presentation bootstrap dimensions or slide count');
  }
  assertNullableString(candidate.defaultTextColor, 'defaultTextColor');
  assertNullableString(candidate.majorFont, 'majorFont');
  assertNullableString(candidate.minorFont, 'minorFont');
  assertNullableString(candidate.hlinkColor, 'hlinkColor');
  assertNullableString(candidate.folHlinkColor, 'folHlinkColor');
  return Object.freeze({
    slideCount: candidate.slideCount as number,
    slideWidth: candidate.slideWidth as number,
    slideHeight: candidate.slideHeight as number,
    defaultTextColor: candidate.defaultTextColor,
    majorFont: candidate.majorFont,
    minorFont: candidate.minorFont,
    hlinkColor: candidate.hlinkColor,
    folHlinkColor: candidate.folHlinkColor,
    slides: Object.freeze(candidate.slides.map(copyBootstrapSlide)),
  });
}

function copyMediaElement(element: MediaElement): Readonly<MediaElement> {
  return Object.freeze({
    type: 'media',
    x: element.x,
    y: element.y,
    width: element.width,
    height: element.height,
    rotation: element.rotation,
    flipH: element.flipH,
    flipV: element.flipV,
    mediaKind: element.mediaKind,
    posterPath: element.posterPath,
    posterMimeType: element.posterMimeType,
    mediaPath: element.mediaPath,
    mimeType: element.mimeType,
  });
}

function normalizeMediaElement(value: unknown, slideIndex: number): Readonly<MediaElement> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`invalid PPTX presentation preflight media at slide ${slideIndex}`);
  }
  const element = value as Partial<MediaElement>;
  for (const field of ['x', 'y', 'width', 'height', 'rotation'] as const) {
    if (typeof element[field] !== 'number' || !Number.isFinite(element[field])) {
      throw new Error(`invalid PPTX presentation preflight media ${field} at slide ${slideIndex}`);
    }
  }
  if (
    element.type !== 'media' ||
    typeof element.flipH !== 'boolean' ||
    typeof element.flipV !== 'boolean' ||
    (element.mediaKind !== 'audio' && element.mediaKind !== 'video') ||
    typeof element.posterPath !== 'string' ||
    typeof element.posterMimeType !== 'string' ||
    typeof element.mediaPath !== 'string' ||
    typeof element.mimeType !== 'string'
  ) {
    throw new Error(`invalid PPTX presentation preflight media fields at slide ${slideIndex}`);
  }
  return copyMediaElement(element as MediaElement);
}

/** Validate, detach, and freeze a compact preflight crossing a worker boundary. */
export function normalizePresentationPreflight(value: unknown): PresentationPreflight {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error('invalid PPTX presentation preflight payload');
  }
  const candidate = value as Partial<PresentationPreflight>;
  if (
    !Number.isSafeInteger(candidate.slideCount) || (candidate.slideCount ?? -1) < 0 ||
    !Number.isSafeInteger(candidate.slideWidth) || (candidate.slideWidth ?? 0) <= 0 ||
    !Number.isSafeInteger(candidate.slideHeight) || (candidate.slideHeight ?? 0) <= 0 ||
    !Array.isArray(candidate.slides) ||
    candidate.slides.length !== candidate.slideCount ||
    !Array.isArray(candidate.fontPreloadNames)
  ) {
    throw new Error('invalid PPTX presentation preflight dimensions or slide count');
  }
  assertNullableString(candidate.defaultTextColor, 'defaultTextColor');
  assertNullableString(candidate.majorFont, 'majorFont');
  assertNullableString(candidate.minorFont, 'minorFont');
  assertNullableString(candidate.hlinkColor, 'hlinkColor');
  assertNullableString(candidate.folHlinkColor, 'folHlinkColor');
  const slides = candidate.slides.map((value, index): PresentationPreflightSlide => {
    if (!value || typeof value !== 'object' || Array.isArray(value)) {
      throw new Error(`invalid PPTX presentation preflight slide at ${index}`);
    }
    const slide = value as Partial<PresentationPreflightSlide>;
    if (
      slide.index !== index ||
      (slide.partName !== undefined && typeof slide.partName !== 'string') ||
      (slide.notes !== null && typeof slide.notes !== 'string') ||
      typeof slide.hidden !== 'boolean' ||
      !Array.isArray(slide.mediaElements)
    ) {
      throw new Error(`invalid PPTX presentation preflight slide fields at ${index}`);
    }
    return Object.freeze({
      index,
      ...(slide.partName === undefined ? {} : { partName: slide.partName }),
      notes: slide.notes,
      hidden: slide.hidden,
      mediaElements: Object.freeze(
        slide.mediaElements.map((media) => normalizeMediaElement(media, index)),
      ),
    });
  });
  const fontPreloadNames = candidate.fontPreloadNames.map((name, index) => {
    if (name !== null && typeof name !== 'string') {
      throw new Error(`invalid PPTX presentation preflight font at ${index}`);
    }
    return name;
  });
  return Object.freeze({
    slideCount: candidate.slideCount as number,
    slideWidth: candidate.slideWidth as number,
    slideHeight: candidate.slideHeight as number,
    defaultTextColor: candidate.defaultTextColor,
    majorFont: candidate.majorFont,
    minorFont: candidate.minorFont,
    hlinkColor: candidate.hlinkColor,
    folHlinkColor: candidate.folHlinkColor,
    slides: Object.freeze(slides),
    fontPreloadNames: Object.freeze(fontPreloadNames),
  });
}

export function findPreflightMimeType(
  preflight: PresentationPreflight,
  partPath: string,
): string {
  for (const slide of preflight.slides) {
    for (const media of slide.mediaElements) {
      if (media.mediaPath === partPath) return media.mimeType;
      if (media.posterPath === partPath) return media.posterMimeType;
    }
  }
  return '';
}

function projectSlide(
  slide: Slide,
  descriptor: PresentationBootstrapSlide,
): PresentationPreflightSlide {
  if (slide.index !== descriptor.index || slide.partName !== descriptor.partName) {
    throw new Error(`PPTX pulled slide identity does not match bootstrap index ${descriptor.index}`);
  }
  return Object.freeze({
    index: descriptor.index,
    ...(descriptor.partName === undefined ? {} : { partName: descriptor.partName }),
    notes: slide.notes ?? null,
    hidden: slide.hidden ?? false,
    mediaElements: Object.freeze(
      slide.elements
        .filter((element): element is MediaElement => element.type === 'media')
        .map(copyMediaElement),
    ),
  });
}

export function assertPresentationPreflightProjectionBytes(
  observed: number,
  usage: OoxmlResourceUsageSnapshot = ZERO_RESOURCE_USAGE,
): void {
  assertProjectionBytes(observed, PPTX_MAX_PREFLIGHT_PROJECTION_BYTES, usage);
}

function assertProjectionBytes(
  observed: number,
  limit: number,
  usage: OoxmlResourceUsageSnapshot,
): void {
  if (observed <= limit) return;
  throw new OoxmlResourceLimitError(
    `PPTX presentation preflight exceeded its hard limit of ${limit} projected bytes`,
    {
      stage: 'parsing',
      violation: {
        format: 'pptx',
        operation: 'presentation-preflight',
        resource: 'presentation-preflight',
        metric: 'projected-bytes',
        limit,
        observed: Math.min(observed, limit + 1),
        configurable: false,
        usage,
      },
    },
  );
}

/** Transaction returned directly from a SlidePullWorker `acceptSlide` hook. */
export interface PreparedPresentationPreflightSlide {
  /** Conservative retained-state projection while the candidate awaits ACK. */
  readonly projectedBytes: number;
  rollback(): void;
  commit(): void;
}

interface PendingAcceptance {
  state: 'prepared' | 'committed' | 'rolled-back';
  readonly fact: PresentationPreflightSlide;
  readonly fonts: PptxFontPreloadAccumulator;
  readonly fontNames: readonly (string | null)[];
  readonly fontBytes: number;
  readonly committedBytes: number;
}

/** @internal Test-only lowering; production cannot raise or replace the hard ceiling. */
export interface PresentationPreflightBuilderOptions {
  readonly hardLimitForTesting?: number;
}

/**
 * Sequential admission builder. It never retains a source Slide: each accepted
 * unit contributes immutable compact facts and script flags, then can be ACKed
 * and released by the pull-session owner.
 */
export class PresentationPreflightBuilder {
  private readonly slideCountValue: number;
  private readonly slideWidthValue: number;
  private readonly slideHeightValue: number;
  private readonly defaultTextColorValue: string | null;
  private readonly majorFontValue: string | null;
  private readonly minorFontValue: string | null;
  private readonly hlinkColorValue: string | null;
  private readonly folHlinkColorValue: string | null;
  private descriptors: (PresentationBootstrapSlide | undefined)[];
  private slides: PresentationPreflightSlide[] = [];
  private fonts: PptxFontPreloadAccumulator;
  private fontPreloadNames: readonly (string | null)[];
  private fontProjectionBytes: number;
  private projectionBytesValue: number;
  private readonly limit: number;
  private pending: PendingAcceptance | null = null;
  private finished: PresentationPreflight | null = null;

  constructor(
    bootstrap: PresentationBootstrap,
    options: PresentationPreflightBuilderOptions = {},
  ) {
    const normalized = normalizePresentationBootstrap(bootstrap);
    const requestedLimit = options.hardLimitForTesting ?? PPTX_MAX_PREFLIGHT_PROJECTION_BYTES;
    if (
      !Number.isSafeInteger(requestedLimit) || requestedLimit <= 0 ||
      requestedLimit > PPTX_MAX_PREFLIGHT_PROJECTION_BYTES
    ) {
      throw new Error('invalid PPTX presentation preflight test limit');
    }
    this.limit = requestedLimit;
    this.slideCountValue = normalized.slideCount;
    this.slideWidthValue = normalized.slideWidth;
    this.slideHeightValue = normalized.slideHeight;
    this.defaultTextColorValue = normalized.defaultTextColor;
    this.majorFontValue = normalized.majorFont;
    this.minorFontValue = normalized.minorFont;
    this.hlinkColorValue = normalized.hlinkColor;
    this.folHlinkColorValue = normalized.folHlinkColor;
    this.descriptors = [...normalized.slides];
    this.fonts = new PptxFontPreloadAccumulator(
      this.majorFontValue,
      this.minorFontValue,
    );
    this.fontPreloadNames = Object.freeze(this.fonts.names());
    this.fontProjectionBytes = measureStructuralJson(
      this.fontPreloadNames,
      this.limit,
    ).jsonBytes;
    this.projectionBytesValue = measureStructuralJson({
      slideCount: this.slideCountValue,
      slideWidth: this.slideWidthValue,
      slideHeight: this.slideHeightValue,
      defaultTextColor: this.defaultTextColorValue,
      majorFont: this.majorFontValue,
      minorFont: this.minorFontValue,
      hlinkColor: this.hlinkColorValue,
      folHlinkColor: this.folHlinkColorValue,
      remainingSlides: this.descriptors,
      slides: [],
      fontPreloadNames: this.fontPreloadNames,
    }, this.limit).jsonBytes;
    assertProjectionBytes(this.projectionBytesValue, this.limit, ZERO_RESOURCE_USAGE);
  }

  get acceptedSlideCount(): number {
    return this.finished?.slideCount ?? this.slides.length;
  }

  get projectedBytes(): number {
    return this.projectionBytesValue;
  }

  get remainingDescriptorCount(): number {
    return this.descriptors.reduce((count, descriptor) => count + Number(descriptor !== undefined), 0);
  }

  addSlide(
    slide: Slide,
    usage: OoxmlResourceUsageSnapshot = ZERO_RESOURCE_USAGE,
  ): void {
    this.prepareSlide(slide, usage).commit();
  }

  prepareSlide(
    slide: Slide,
    usage: OoxmlResourceUsageSnapshot = ZERO_RESOURCE_USAGE,
  ): PreparedPresentationPreflightSlide {
    if (this.finished) throw new Error('PPTX presentation preflight is already finished');
    if (this.pending) throw new Error('PPTX presentation preflight already has a prepared slide');
    const index = this.slides.length;
    const descriptor = this.descriptors[index];
    if (!descriptor) throw new Error('PPTX presentation preflight received an extra slide');
    const fact = projectSlide(slide, descriptor);
    const nextFonts = this.fonts.withSlide(slide);
    const nextFontNames = Object.freeze(nextFonts.names());
    const nextFontBytes = measureStructuralJson(
      nextFontNames,
      this.limit,
    ).jsonBytes;
    const slideBytes = measureStructuralJson(
      fact,
      this.limit,
    ).jsonBytes;
    // Commit replaces the current descriptor with JSON `null`, releases its
    // strings, and adds one fact. This exact delta keeps the building-state
    // projection honest while descriptors and committed facts coexist.
    let committedBytes = this.projectionBytesValue
      - this.fontProjectionBytes
      - measureStructuralJson(descriptor, this.limit).jsonBytes
      + 4;
    committedBytes = cappedAdd(committedBytes, nextFontBytes, this.limit);
    committedBytes = cappedAdd(committedBytes, slideBytes, this.limit);
    if (this.slides.length !== 0) {
      committedBytes = cappedAdd(committedBytes, 1, this.limit);
    }
    // Before Rust ACK, both the unchanged builder and candidate acceptance
    // state are live. Charge the complete candidate projection in addition to
    // the existing retained state, then enforce the worse of prepare/commit.
    const candidateBytes = measureStructuralJson({
      slide: fact,
      fontPreloadNames: nextFontNames,
    }, this.limit).jsonBytes;
    const preparedBytes = cappedAdd(this.projectionBytesValue, candidateBytes, this.limit);
    const observed = Math.max(preparedBytes, committedBytes);
    assertProjectionBytes(observed, this.limit, usage);
    const pending: PendingAcceptance = {
      state: 'prepared',
      fact,
      fonts: nextFonts,
      fontNames: nextFontNames,
      fontBytes: nextFontBytes,
      committedBytes,
    };
    this.pending = pending;
    return {
      projectedBytes: preparedBytes,
      commit: () => {
        if (pending.state === 'committed') return;
        if (pending.state === 'rolled-back') {
          throw new Error('PPTX presentation preflight cannot commit a rolled-back slide');
        }
        if (this.pending !== pending) {
          throw new Error('PPTX presentation preflight prepared slide is stale');
        }
        this.descriptors[index] = undefined;
        this.slides.push(pending.fact);
        this.fonts = pending.fonts;
        this.fontPreloadNames = pending.fontNames;
        this.fontProjectionBytes = pending.fontBytes;
        this.projectionBytesValue = pending.committedBytes;
        pending.state = 'committed';
        this.pending = null;
      },
      rollback: () => {
        if (pending.state === 'rolled-back') return;
        if (pending.state === 'committed') {
          throw new Error('PPTX presentation preflight cannot roll back a committed slide');
        }
        if (this.pending !== pending) {
          throw new Error('PPTX presentation preflight prepared slide is stale');
        }
        pending.state = 'rolled-back';
        this.pending = null;
      },
    };
  }

  finish(): PresentationPreflight {
    if (this.finished) return this.finished;
    if (this.pending) throw new Error('PPTX presentation preflight has an uncommitted slide');
    if (this.slides.length !== this.slideCountValue) {
      throw new Error(
        `PPTX presentation preflight is incomplete: ${this.slides.length}/${this.slideCountValue} slides`,
      );
    }
    this.finished = Object.freeze({
      slideCount: this.slideCountValue,
      slideWidth: this.slideWidthValue,
      slideHeight: this.slideHeightValue,
      defaultTextColor: this.defaultTextColorValue,
      majorFont: this.majorFontValue,
      minorFont: this.minorFontValue,
      hlinkColor: this.hlinkColorValue,
      folHlinkColor: this.folHlinkColorValue,
      slides: Object.freeze([...this.slides]),
      fontPreloadNames: this.fontPreloadNames,
    });
    // The frozen compact model owns its slide-array storage from here. The
    // builder releases both construction-only arrays rather than retaining a
    // second array of fact references or descriptor slots after finish.
    this.descriptors = [];
    this.slides = [];
    this.projectionBytesValue = measureStructuralJson(
      this.finished,
      this.limit,
    ).jsonBytes;
    return this.finished;
  }
}
