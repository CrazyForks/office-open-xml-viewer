// Compatibility facade for legacy DOCX modules. The layout-owned implementation
// lives under layout/ so canonical layout code never depends on a root legacy
// module while existing internal import paths remain stable during migration.
export {
  FLOAT_OVERLAP_EPS,
  FLOAT_PAGE_RIGHT_SLACK,
  LINE_START_GAP_EPS_PT,
  MIN_LINE_GAP,
  WORD_MIN_LINE_START_PT,
  floatOverlapsColumnX,
  isWrapFloat,
  normalizeWrapSide,
  prepareFloatWrap,
  rectsOverlap,
  resolveFloatOverlap,
  resolveLineFloatWindow,
  computePreparedLineFloatWindow,
  computePreparedLineFloatWindowWithDiagnostics,
  skipPastTopAndBottom,
  widestFreeGap,
  wordMinLineStartPx,
} from './layout/float-wrap.js';

export type {
  FloatRect,
  Gap,
  LineFloatReference,
  LineFloatSweepDiagnostics,
  PreparedFloatWrap,
  WrapSide,
} from './layout/float-wrap.js';
