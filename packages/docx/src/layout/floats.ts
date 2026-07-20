import {
  WORD_FLOAT_DIFFERENT_PARAGRAPH_DISPLACEMENT,
} from './compatibility.js';
import {
  resolveAxisAlignedOverlap,
  type AxisAlignedRect,
} from './axis-aligned-overlap.js';
import type { LayoutRect } from './types.js';

/** Numerical tolerances retained by callers while C1b removes the temporary
 * scalar registry. Point-space table transactions also used these exact values
 * through that adapter before C1a and must preserve them explicitly. */
export const FLOAT_OVERLAP_EPS = 0.01;
export const FLOAT_PAGE_RIGHT_SLACK = 0.5;

export type FloatPlacementKind = 'drawingml' | 'table' | 'frame';

/** Immutable object and text-exclusion geometry in one caller-selected
 * coordinate space. Normative object collisions use `bounds`; Word-compatible
 * text-wrap displacement uses `exclusionBounds`. */
export interface FloatPlacementParticipant {
  readonly occurrenceId: string;
  readonly kind: FloatPlacementKind;
  readonly paragraphId: number;
  readonly bounds: LayoutRect;
  readonly exclusionBounds: LayoutRect;
}

export type FloatAvoidance =
  | Readonly<{ kind: 'drawingml-normative' }>
  | Readonly<{ kind: 'floating-table-never' }>
  | Readonly<{ kind: 'word-different-paragraph'; paragraphId: number }>
  | Readonly<{ kind: 'none' }>;

export interface FloatPlacement {
  readonly bounds: LayoutRect;
  readonly exclusionBounds: LayoutRect;
  readonly displacement: Readonly<{ xPt: number; yPt: number }>;
  readonly appliedCompatibilityRuleIds: readonly string[];
}

export interface ResolveFloatPlacementInput {
  readonly moving: FloatPlacementParticipant;
  readonly blockers: readonly FloatPlacementParticipant[];
  readonly avoidance: FloatAvoidance;
  readonly rightBoundaryPt: number;
  /** Legacy display-space callers may temporarily supply their existing
   * numerical tolerance. Retained point-space callers omit both values. */
  readonly overlapEpsilonPt?: number;
  readonly rightBoundarySlackPt?: number;
}

function axisRect(rect: LayoutRect): AxisAlignedRect {
  return {
    left: rect.xPt,
    right: rect.xPt + rect.widthPt,
    top: rect.yPt,
    bottom: rect.yPt + rect.heightPt,
  };
}

function translateRect(
  rect: LayoutRect,
  xPt: number,
  yPt: number,
): LayoutRect {
  if (xPt === 0 && yPt === 0) return rect;
  return Object.freeze({
    xPt: rect.xPt + xPt,
    yPt: rect.yPt + yPt,
    widthPt: rect.widthPt,
    heightPt: rect.heightPt,
  });
}

function placement(
  moving: FloatPlacementParticipant,
  xPt: number,
  yPt: number,
  ruleIds: readonly string[],
): FloatPlacement {
  return Object.freeze({
    bounds: translateRect(moving.bounds, xPt, yPt),
    exclusionBounds: translateRect(moving.exclusionBounds, xPt, yPt),
    displacement: Object.freeze({ xPt, yPt }),
    appliedCompatibilityRuleIds: Object.freeze([...ruleIds]),
  });
}

/**
 * Resolve only float-to-float displacement policy.
 *
 * Anchor axis/size/wrap geometry remains owned by `anchor-frame.ts`; text-line
 * exclusion remains owned by `float-wrap.ts`; page/column admission remains
 * owned by the paginator. ECMA-376 chooses the eligible object class but does
 * not choose a valid displacement direction, so the shared deterministic
 * right-then-down kernel preserves the established direction.
 */
export function resolveFloatPlacement(
  input: ResolveFloatPlacementInput,
): FloatPlacement {
  const { moving, avoidance } = input;
  if (avoidance.kind === 'none') return placement(moving, 0, 0, []);

  const compatibility = avoidance.kind === 'word-different-paragraph';
  const movingRect = compatibility ? moving.exclusionBounds : moving.bounds;
  const eligible = input.blockers.filter((blocker) => {
    if (avoidance.kind === 'drawingml-normative') {
      return blocker.kind === 'drawingml';
    }
    if (avoidance.kind === 'floating-table-never') {
      return blocker.kind === 'table';
    }
    return blocker.paragraphId !== avoidance.paragraphId;
  });
  if (eligible.length === 0) {
    return placement(moving, 0, 0, []);
  }
  const resolved = resolveAxisAlignedOverlap(
    axisRect(movingRect),
    eligible.map((blocker) => axisRect(
      compatibility ? blocker.exclusionBounds : blocker.bounds,
    )),
    {
      overlapEpsilon: input.overlapEpsilonPt ?? 0,
      rightBoundary: input.rightBoundaryPt,
      rightBoundarySlack: input.rightBoundarySlackPt ?? 0,
    },
  );
  const xPt = resolved.left - movingRect.xPt;
  const yPt = resolved.top - movingRect.yPt;
  return placement(moving, xPt, yPt, compatibility && (xPt !== 0 || yPt !== 0)
    ? [WORD_FLOAT_DIFFERENT_PARAGRAPH_DISPLACEMENT.id]
    : []);
}

interface ScalarFloatParticipant {
  readonly kind: 'table' | 'shape' | 'frame';
  readonly xLeft: number;
  readonly xRight: number;
  readonly yTop: number;
  readonly yBottom: number;
  readonly paraId: number;
  readonly imageX?: number;
  readonly imageY?: number;
  readonly imageW?: number;
  readonly imageH?: number;
}

function scalarParticipant(
  float: ScalarFloatParticipant,
  index: number,
): FloatPlacementParticipant {
  const xLeft = float.xLeft;
  const xRight = float.xRight;
  const yTop = float.yTop;
  const yBottom = float.yBottom;
  const hasObjectBounds = [
    float.imageX, float.imageY, float.imageW, float.imageH,
  ].every((value) => typeof value === 'number' && Number.isFinite(value));
  const exclusionBounds = {
    xPt: xLeft,
    yPt: yTop,
    widthPt: xRight - xLeft,
    heightPt: yBottom - yTop,
  };
  return {
    occurrenceId: `legacy-float:${index}`,
    kind: float.kind === 'shape' ? 'drawingml' : float.kind,
    paragraphId: float.paraId,
    bounds: hasObjectBounds ? {
      xPt: float.imageX!,
      yPt: float.imageY!,
      widthPt: float.imageW!,
      heightPt: float.imageH!,
    } : exclusionBounds,
    exclusionBounds,
  };
}

/**
 * Temporary compatibility signature for root renderer helpers. All blocker
 * scoping and displacement now delegates to `resolveFloatPlacement`; C1b
 * removes this scalar adapter when the display-space registry is retired.
 */
export function resolveFloatOverlap(
  x: number, y: number, w: number, h: number,
  dl: number, dr: number, dt: number, db: number,
  paraId: number, allowOverlap: boolean,
  kind: ScalarFloatParticipant['kind'],
  pageRight: number,
  floats: readonly ScalarFloatParticipant[],
): { x: number; y: number } {
  const moving: FloatPlacementParticipant = {
    occurrenceId: 'legacy-moving-float',
    kind: kind === 'shape' ? 'drawingml' : kind,
    paragraphId: paraId,
    bounds: { xPt: x, yPt: y, widthPt: w, heightPt: h },
    exclusionBounds: {
      xPt: x - dl,
      yPt: y - dt,
      widthPt: w + dl + dr,
      heightPt: h + dt + db,
    },
  };
  try {
    const result = resolveFloatPlacement({
      moving,
      blockers: floats.map(scalarParticipant),
      avoidance: allowOverlap
        ? { kind: 'word-different-paragraph', paragraphId: paraId }
        : kind === 'table'
          ? { kind: 'floating-table-never' }
          : { kind: 'drawingml-normative' },
      rightBoundaryPt: pageRight,
      overlapEpsilonPt: FLOAT_OVERLAP_EPS,
      rightBoundarySlackPt: FLOAT_PAGE_RIGHT_SLACK,
    });
    return { x: result.bounds.xPt, y: result.bounds.yPt };
  } catch (error) {
    if (
      error instanceof Error
      && error.message === 'Axis-aligned overlap resolution did not converge'
    ) {
      throw new Error('Float overlap resolution did not converge');
    }
    throw error;
  }
}
