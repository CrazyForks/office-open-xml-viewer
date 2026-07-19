import { createExactStateTracker, observeExactState } from './repeated-state.js';

interface LineWrapStateSegment {
  readonly src?: unknown;
  readonly text?: string;
}

export interface LineWrapStateLine {
  readonly consumedEnd?: unknown;
  readonly topY?: number;
  readonly xOffset?: number;
  readonly availWidth?: number;
  readonly segments: readonly LineWrapStateSegment[];
}

export class LineWrapNonConvergenceError extends Error {
  readonly code = 'line-wrap-non-convergence';

  constructor(readonly states: readonly string[]) {
    super(`Line wrap measure/resolve cycle did not converge (${states.length} states)`);
    this.name = 'LineWrapNonConvergenceError';
  }
}

export function cloneSegmentsForLinePass<T extends object>(segments: readonly T[]): T[] {
  return segments.map((segment) => ({ ...segment }));
}

function lineWrapState(
  lines: readonly LineWrapStateLine[],
  probeHeights: readonly number[],
): string {
  return JSON.stringify(lines.map((line, index) => ({
    end: line.consumedEnd,
    topY: line.topY,
    xOffset: line.xOffset,
    availableWidth: line.availWidth,
    probeHeight: probeHeights[index],
    segments: line.segments.map((segment) => ({
      source: segment.src,
      ...(segment.text === undefined ? {} : { text: segment.text }),
    })),
  })));
}

/**
 * Line boundaries, polygon windows, and line-box heights each select from
 * finite exact state sets. An adjacent repeat is therefore the fixed point;
 * a non-adjacent repeat is a real cycle and must remain observable rather than
 * being hidden behind an empirical pass limit.
 */
export function convergeLineWrap<TLine extends LineWrapStateLine>(
  measure: (probeHeights: readonly number[] | null) => TLine[],
  lineBoxHeight: (line: TLine) => number,
): TLine[] {
  const tracker = createExactStateTracker();
  let probes: readonly number[] | null = null;
  for (;;) {
    const lines = measure(probes);
    const nextProbes = Object.freeze(lines.map(lineBoxHeight));
    const observation = observeExactState(tracker, lineWrapState(lines, nextProbes));
    if (observation.kind === 'fixed') return lines;
    if (observation.kind === 'cycle') {
      throw new LineWrapNonConvergenceError(observation.states);
    }
    probes = nextProbes;
  }
}
