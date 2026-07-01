import { describe, it, expect } from 'vitest';
import type { XlsxViewer, XlsxViewerOptions, HiddenSheetMode } from './viewer.js';
import { nextVisibleIndex, resolveVisibleIndex } from '@silurus/ooxml-core';

/**
 * Compile-time API-surface assertions (erased at runtime, enforced by
 * `pnpm typecheck`). XlsxViewer is DOM-bound and the vitest env is `node`, so —
 * like pptx's viewer-hidden.test.ts — the viewer is verified at the type level;
 * its skip policy is delegated to the pure core helpers tested in core.
 */
type Expect<T extends true> = T;
type Equal<A, B> =
  (<T>() => T extends A ? 1 : 2) extends (<T>() => T extends B ? 1 : 2) ? true : false;

type _Modes = Expect<Equal<HiddenSheetMode, 'show' | 'skip' | 'dim'>>;
type _SetMode = Expect<Equal<XlsxViewer['setHiddenSheetMode'], (mode: HiddenSheetMode) => Promise<void>>>;
type _ModeGetter = Expect<Equal<XlsxViewer['hiddenSheetMode'], HiddenSheetMode>>;
type _VisibleCount = Expect<Equal<XlsxViewer['visibleSheetCount'], number>>;
const _opts: XlsxViewerOptions = { hiddenSheetMode: 'skip' };

describe('XlsxViewer hidden-sheet policy (delegated core helpers)', () => {
  it('skip navigation jumps over hidden sheets via the tested core helpers', () => {
    const isHidden = (i: number) => i === 1; // visible: 0, 2
    expect(nextVisibleIndex(0, 1, isHidden, 3)).toBe(2);
    expect(resolveVisibleIndex(1, isHidden, 3)).toBe(2);
    void _opts;
  });
});
