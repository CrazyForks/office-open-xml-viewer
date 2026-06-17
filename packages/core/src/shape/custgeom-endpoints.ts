import type { PathCmd } from '../types/common';

/**
 * An endpoint of a custom-geometry path, used to place a line-end decoration
 * (arrow head) on a freeform / curved line.
 *
 * Coordinates `(x, y)` are in the same **normalised** [0,1] space as the
 * incoming {@link PathCmd}s (relative to the shape bounding box). `(dx, dy)` is
 * the **outward** tangent direction at that endpoint, also expressed in
 * normalised space — i.e. it points *away* from the line so an arrow head drawn
 * along it faces outward (head at the start, tail at the end).
 *
 * The tangent is returned as a raw direction vector rather than a baked angle
 * because the shape may be scaled anisotropically (w ≠ h). The caller scales the
 * vector by the box dimensions — `atan2(dy·h, dx·w)` — to obtain the device-space
 * orientation. {@link CustGeomEndpoint.angle} provides the isotropic angle
 * (`atan2(dy, dx)`) for convenience / testing.
 */
export interface CustGeomEndpoint {
  /** Normalised x in [0,1]. */
  x: number;
  /** Normalised y in [0,1]. */
  y: number;
  /** Outward tangent x-component (normalised space). */
  dx: number;
  /** Outward tangent y-component (normalised space). */
  dy: number;
  /** Isotropic outward angle in radians (`atan2(dy, dx)`). */
  angle: number;
}

/** Two endpoints may be `null` independently (closed end ⇒ no arrow head). */
export interface CustGeomEndpoints {
  start: CustGeomEndpoint | null;
  end: CustGeomEndpoint | null;
}

const EPS = 1e-9;

function isDrawCmd(cmd: PathCmd): boolean {
  return cmd.cmd === 'lineTo' || cmd.cmd === 'cubicBezTo' || cmd.cmd === 'arcTo';
}

function makeEndpoint(x: number, y: number, dx: number, dy: number): CustGeomEndpoint {
  // Normalise signed zero so the tangent (and any atan2 derived from it) is
  // canonical: `atan2(-0, -k)` is -π while `atan2(0, -k)` is +π — same
  // direction, but the negative zero leaks an arbitrary sign into the angle.
  const ndx = dx === 0 ? 0 : dx;
  const ndy = dy === 0 ? 0 : dy;
  return { x, y, dx: ndx, dy: ndy, angle: Math.atan2(ndy, ndx) };
}

/**
 * Forward tangent (direction of travel) at the **start** of a sub-path, i.e.
 * leaving `(penX, penY)` along the first drawn command `cmd`.
 */
function startForwardTangent(
  penX: number,
  penY: number,
  cmd: PathCmd,
): { dx: number; dy: number } {
  switch (cmd.cmd) {
    case 'lineTo':
      return { dx: cmd.x - penX, dy: cmd.y - penY };
    case 'cubicBezTo': {
      // Tangent at t=0 of a cubic Bézier points toward the first control point.
      let dx = cmd.x1 - penX;
      let dy = cmd.y1 - penY;
      if (Math.abs(dx) < EPS && Math.abs(dy) < EPS) {
        // Degenerate first control point: fall back to the second, then the end.
        dx = cmd.x2 - penX;
        dy = cmd.y2 - penY;
      }
      if (Math.abs(dx) < EPS && Math.abs(dy) < EPS) {
        dx = cmd.x - penX;
        dy = cmd.y - penY;
      }
      return { dx, dy };
    }
    case 'arcTo': {
      // Ellipse: P(t) = (cx + wr cos t, cy + hr sin t), with the pen at t=stAng
      // (matches buildCustomPath). Tangent dP/dt = (-wr sin t, hr cos t); for
      // sweep>0 travel increases t, for sweep<0 it decreases t.
      const stRad = (cmd.stAng * Math.PI) / 180;
      const dir = cmd.swAng < 0 ? -1 : 1;
      const dx = -cmd.wr * Math.sin(stRad) * dir;
      const dy = cmd.hr * Math.cos(stRad) * dir;
      return { dx, dy };
    }
    default:
      return { dx: 0, dy: 0 };
  }
}

/**
 * Forward tangent (direction of travel) **arriving** at the end of a drawn
 * command, given the point `(prevX, prevY)` it started from.
 */
function endForwardTangent(
  prevX: number,
  prevY: number,
  cmd: PathCmd,
): { dx: number; dy: number; x: number; y: number } {
  switch (cmd.cmd) {
    case 'lineTo':
      return { dx: cmd.x - prevX, dy: cmd.y - prevY, x: cmd.x, y: cmd.y };
    case 'cubicBezTo': {
      // Tangent at t=1 points from the second control point to the end.
      let dx = cmd.x - cmd.x2;
      let dy = cmd.y - cmd.y2;
      if (Math.abs(dx) < EPS && Math.abs(dy) < EPS) {
        // Degenerate: second control coincides with the end → use first control.
        dx = cmd.x - cmd.x1;
        dy = cmd.y - cmd.y1;
      }
      if (Math.abs(dx) < EPS && Math.abs(dy) < EPS) {
        // Fully degenerate → chord from the previous point to the end.
        dx = cmd.x - prevX;
        dy = cmd.y - prevY;
      }
      return { dx, dy, x: cmd.x, y: cmd.y };
    }
    case 'arcTo': {
      const stRad = (cmd.stAng * Math.PI) / 180;
      const swRad = (cmd.swAng * Math.PI) / 180;
      const endRad = stRad + swRad;
      // Centre back-calculated from the pen at stAng (matches buildCustomPath):
      // cx = penX - wr cos(stAng); end = (cx + wr cos(endRad), cy + hr sin(endRad)).
      const cxOff = -cmd.wr * Math.cos(stRad); // centre relative to prev point
      const cyOff = -cmd.hr * Math.sin(stRad);
      const ex = prevX + cxOff + cmd.wr * Math.cos(endRad);
      const ey = prevY + cyOff + cmd.hr * Math.sin(endRad);
      const dir = cmd.swAng < 0 ? -1 : 1;
      const dx = -cmd.wr * Math.sin(endRad) * dir;
      const dy = cmd.hr * Math.cos(endRad) * dir;
      return { dx, dy, x: ex, y: ey };
    }
    default:
      return { dx: 0, dy: 0, x: prevX, y: prevY };
  }
}

/** Track the pen position across one sub-path to find its final terminal point. */
function terminalPoint(cmds: PathCmd[]): { x: number; y: number } | null {
  let px = 0;
  let py = 0;
  let started = false;
  for (const cmd of cmds) {
    switch (cmd.cmd) {
      case 'moveTo':
        px = cmd.x; py = cmd.y; started = true;
        break;
      case 'lineTo':
      case 'cubicBezTo':
        px = cmd.x; py = cmd.y;
        break;
      case 'arcTo': {
        const stRad = (cmd.stAng * Math.PI) / 180;
        const endRad = stRad + (cmd.swAng * Math.PI) / 180;
        const cx = px - cmd.wr * Math.cos(stRad);
        const cy = py - cmd.hr * Math.sin(stRad);
        px = cx + cmd.wr * Math.cos(endRad);
        py = cy + cmd.hr * Math.sin(endRad);
        break;
      }
      case 'close':
        break;
    }
  }
  return started ? { x: px, y: py } : null;
}

/** Does this sub-path form a closed loop (explicit `close` or terminal ≈ start)? */
function isClosed(cmds: PathCmd[]): boolean {
  if (cmds.some((c) => c.cmd === 'close')) return true;
  const first = cmds.find((c) => c.cmd === 'moveTo') as
    | Extract<PathCmd, { cmd: 'moveTo' }>
    | undefined;
  if (!first) return false;
  const term = terminalPoint(cmds);
  if (!term) return false;
  // Only a path that actually draws something can be an implicit loop.
  const hasDraw = cmds.some(isDrawCmd);
  if (!hasDraw) return false;
  return Math.abs(term.x - first.x) < EPS && Math.abs(term.y - first.y) < EPS;
}

/**
 * Extract the start and end decoration points of a custom-geometry path.
 *
 * - `start` = the very first `moveTo` of the first sub-path. Its tangent is the
 *   **reverse** of the direction leaving that point (so a head faces outward).
 * - `end`   = the terminal point of the last drawn command of the last sub-path.
 *   Its tangent is the direction of travel arriving there (already outward).
 *
 * An end belonging to a **closed** sub-path returns `null` (no arrow head), which
 * matches PowerPoint: line-end decorations only apply to open paths.
 *
 * @param subpaths Normalised (`[0,1]`) custGeom sub-paths.
 */
export function getCustGeomEndpoints(subpaths: PathCmd[][]): CustGeomEndpoints {
  const result: CustGeomEndpoints = { start: null, end: null };
  if (!subpaths || subpaths.length === 0) return result;

  // ── start: first sub-path's opening moveTo + first drawn command ──────────
  const firstSub = subpaths[0];
  if (firstSub && firstSub.length > 0 && !isClosed(firstSub)) {
    const mv = firstSub.find((c) => c.cmd === 'moveTo') as
      | Extract<PathCmd, { cmd: 'moveTo' }>
      | undefined;
    const firstDraw = firstSub.find(isDrawCmd);
    if (mv && firstDraw) {
      const fwd = startForwardTangent(mv.x, mv.y, firstDraw);
      if (Math.abs(fwd.dx) > EPS || Math.abs(fwd.dy) > EPS) {
        // Outward (head) tangent is the reverse of the travel direction.
        result.start = makeEndpoint(mv.x, mv.y, -fwd.dx, -fwd.dy);
      }
    }
  }

  // ── end: last sub-path's final drawn command ─────────────────────────────
  const lastSub = subpaths[subpaths.length - 1];
  if (lastSub && lastSub.length > 0 && !isClosed(lastSub)) {
    // Replay the pen to the point just before the last drawn command.
    let px = 0;
    let py = 0;
    let lastDrawIdx = -1;
    for (let i = 0; i < lastSub.length; i++) {
      if (isDrawCmd(lastSub[i])) lastDrawIdx = i;
    }
    if (lastDrawIdx >= 0) {
      for (let i = 0; i < lastDrawIdx; i++) {
        const cmd = lastSub[i];
        switch (cmd.cmd) {
          case 'moveTo':
          case 'lineTo':
          case 'cubicBezTo':
            px = cmd.x; py = cmd.y;
            break;
          case 'arcTo': {
            const stRad = (cmd.stAng * Math.PI) / 180;
            const endRad = stRad + (cmd.swAng * Math.PI) / 180;
            const cx = px - cmd.wr * Math.cos(stRad);
            const cy = py - cmd.hr * Math.sin(stRad);
            px = cx + cmd.wr * Math.cos(endRad);
            py = cy + cmd.hr * Math.sin(endRad);
            break;
          }
          case 'close':
            break;
        }
      }
      const t = endForwardTangent(px, py, lastSub[lastDrawIdx]);
      if (Math.abs(t.dx) > EPS || Math.abs(t.dy) > EPS) {
        result.end = makeEndpoint(t.x, t.y, t.dx, t.dy);
      }
    }
  }

  return result;
}
