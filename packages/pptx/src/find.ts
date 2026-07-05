/**
 * IX2 pptx find-in-presentation controller.
 *
 * The pptx twin of the docx `DocxFindController`: owns per-slide run lists (the
 * `onTextRun` stream), the matches for the current query, and the active-match
 * cursor. All string/index math is core (`buildTextIndex`, `findMatches`,
 * `nextActive`/`prevActive`); this maps each hit to a `{ slide }` location.
 *
 * The viewer supplies `collectSlideRuns(slide)` — render that slide (to an
 * offscreen canvas) and return its `PptxTextRunInfo[]`. The controller caches
 * per slide until `invalidate()`. The displayed slide's runs are fed in from the
 * visible render so highlight geometry matches exactly what was drawn.
 */
import {
  buildTextIndex,
  findMatches,
  nextActive,
  prevActive,
  type FindMatch,
  type FindMatchesOptions,
  type TextMatch,
} from '@silurus/ooxml-core';
import type { PptxTextRunInfo } from './renderer';

/** Where a pptx match lives: its 0-based slide index. */
export interface PptxMatchLocation {
  slide: number;
}

interface PptxResolvedMatch {
  slide: number;
  text: string;
  slices: TextMatch['slices'];
}

export class PptxFindController {
  private _slideRuns = new Map<number, PptxTextRunInfo[]>();
  private _matches: PptxResolvedMatch[] = [];
  private _active = -1;

  constructor(
    private readonly _slideCount: () => number,
    private readonly _collectSlideRuns: (slide: number) => Promise<PptxTextRunInfo[]>,
  ) {}

  /** Drop all cached runs + matches (call on reload). */
  invalidate(): void {
    this._slideRuns.clear();
    this._matches = [];
    this._active = -1;
  }

  /** The runs for a slide, if scanned (used by the highlight overlay for the
   *  displayed slide). */
  slideRuns(slide: number): PptxTextRunInfo[] | undefined {
    return this._slideRuns.get(slide);
  }

  /** Cache a slide's runs captured from the visible render. */
  setSlideRuns(slide: number, runs: PptxTextRunInfo[]): void {
    this._slideRuns.set(slide, runs);
  }

  /** All match slices on a slide, tagged active — the highlight overlay input. */
  slideHighlights(slide: number): { slices: TextMatch['slices']; active: boolean }[] {
    const out: { slices: TextMatch['slices']; active: boolean }[] = [];
    for (let i = 0; i < this._matches.length; i++) {
      const m = this._matches[i];
      if (m.slide === slide) out.push({ slices: m.slices, active: i === this._active });
    }
    return out;
  }

  /** The active match's slide, or null. */
  activeSlide(): number | null {
    const m = this._matches[this._active];
    return m ? m.slide : null;
  }

  /** The public match list for the current query. */
  matches(): FindMatch<PptxMatchLocation>[] {
    return this._matches.map((m, i) => ({
      matchIndex: i,
      text: m.text,
      location: { slide: m.slide },
    }));
  }

  /** Run a fresh query across every slide, resetting the cursor. */
  async find(query: string, opts: FindMatchesOptions = {}): Promise<FindMatch<PptxMatchLocation>[]> {
    this._matches = [];
    this._active = -1;
    if (query.length === 0) return [];

    const slides = this._slideCount();
    for (let slide = 0; slide < slides; slide++) {
      const runs = await this._ensureSlideRuns(slide);
      const index = buildTextIndex(runs);
      for (const tm of findMatches(index, query, opts)) {
        const text = tm.slices
          .map((s) => runs[s.runIndex].text.slice(s.start, s.end))
          .join('');
        this._matches.push({ slide, text, slices: tm.slices });
      }
    }
    return this.matches();
  }

  next(): FindMatch<PptxMatchLocation> | null {
    this._active = nextActive(this._active, this._matches.length);
    return this._activePublic();
  }

  prev(): FindMatch<PptxMatchLocation> | null {
    this._active = prevActive(this._active, this._matches.length);
    return this._activePublic();
  }

  private _activePublic(): FindMatch<PptxMatchLocation> | null {
    const m = this._matches[this._active];
    if (!m) return null;
    return { matchIndex: this._active, text: m.text, location: { slide: m.slide } };
  }

  private async _ensureSlideRuns(slide: number): Promise<PptxTextRunInfo[]> {
    const cached = this._slideRuns.get(slide);
    if (cached) return cached;
    const runs = await this._collectSlideRuns(slide);
    this._slideRuns.set(slide, runs);
    return runs;
  }
}
