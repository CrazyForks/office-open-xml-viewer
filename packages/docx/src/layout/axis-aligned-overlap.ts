export interface AxisAlignedRect {
  readonly left: number;
  readonly right: number;
  readonly top: number;
  readonly bottom: number;
}

export interface AxisAlignedOverlapPolicy {
  readonly overlapEpsilon: number;
  readonly rightBoundary: number;
  readonly rightBoundarySlack: number;
}

export function axisAlignedRectsOverlap(
  left: AxisAlignedRect,
  right: AxisAlignedRect,
  overlapEpsilon: number,
): boolean {
  return left.left < right.right - overlapEpsilon
    && left.right > right.left + overlapEpsilon
    && left.top < right.bottom - overlapEpsilon
    && left.bottom > right.top + overlapEpsilon;
}

/** Re-seat one axis-aligned rectangle without changing its size.
 *
 * ECMA-376 §20.4.2.3 requires collision avoidance when allowOverlap is false,
 * but does not choose among valid positions. Right-then-down preserves the
 * renderer's pre-retained-layout placement policy while the caller controls
 * which blockers are normative versus compatibility-only. */
export function resolveAxisAlignedOverlap(
  moving: AxisAlignedRect,
  blockers: readonly AxisAlignedRect[],
  policy: AxisAlignedOverlapPolicy,
): Readonly<{ left: number; top: number }> {
  const width = moving.right - moving.left;
  const height = moving.bottom - moving.top;
  if (width < 0 || height < 0) throw new RangeError('Overlap rectangle has negative extent');
  let left = moving.left;
  let top = moving.top;
  for (let moveCount = 0; moveCount <= blockers.length; moveCount += 1) {
    const current = {
      left,
      right: left + width,
      top,
      bottom: top + height,
    };
    // Snapshot each blocker once per move in source-coordinate order. Normal
    // retained registries are immutable; the snapshot also makes the
    // non-convergence guard deterministic for adversarial accessor-backed input.
    const currentBlockers = blockers.map((blocker) => ({
      left: blocker.left,
      right: blocker.right,
      top: blocker.top,
      bottom: blocker.bottom,
    }));
    const intersecting = currentBlockers.filter((blocker) =>
      axisAlignedRectsOverlap(current, blocker, policy.overlapEpsilon));
    if (intersecting.length === 0) return Object.freeze({ left, top });
    if (moveCount === blockers.length) {
      throw new Error('Axis-aligned overlap resolution did not converge');
    }
    const nextLeft = Math.max(...intersecting.map((blocker) => blocker.right));
    if (
      nextLeft + width
      <= policy.rightBoundary + policy.rightBoundarySlack
    ) {
      left = nextLeft;
      continue;
    }
    top = Math.max(...intersecting.map((blocker) => blocker.bottom));
  }
  throw new Error('Axis-aligned overlap resolution did not converge');
}
