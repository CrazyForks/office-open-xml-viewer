// From-scratch UAX#9 engine — the swappable implementation behind BidiEngine.
// The rule set lives in ./rules.ts (conformance-tested against Unicode's
// BidiCharacterTest.txt / BidiTest.txt). This file adapts the code-point-space
// core to the string / code-unit surface renderers use.

import type { BidiEngine } from '../engine.js';
import { REMOVED_LEVEL } from '../types.js';
import type { BaseDirection, BidiLevels } from '../types.js';
import { mirror as mirrorGlyph } from '../char-data.js';
import { resolveLevels, reorderByLevels, REMOVED } from './rules.js';

/** Decode a UTF-16 string into code points + how many code units each occupies. */
function decode(text: string): { cps: number[]; units: number[] } {
  const cps: number[] = [];
  const units: number[] = [];
  for (let i = 0; i < text.length; ) {
    const cp = text.codePointAt(i)!;
    const len = cp > 0xffff ? 2 : 1;
    cps.push(cp);
    units.push(len);
    i += len;
  }
  return { cps, units };
}

class Uax9Engine implements BidiEngine {
  computeLevels(
    text: string,
    base: BaseDirection,
  ): { levels: BidiLevels; paragraphLevel: number } {
    const { cps, units } = decode(text);
    const { levels: cpLevels, paragraphLevel } = resolveLevels(cps, base);

    // Expand per-code-point levels back to per-code-unit (both surrogate halves
    // share the level), mapping the REMOVED sentinel to 255.
    const out = new Uint8Array(text.length);
    let u = 0;
    for (let i = 0; i < cpLevels.length; i++) {
      const v = cpLevels[i] === REMOVED ? REMOVED_LEVEL : cpLevels[i];
      for (let k = 0; k < units[i]; k++) out[u++] = v;
    }
    return { levels: out, paragraphLevel };
  }

  reorderVisual(levels: BidiLevels, start: number, end: number): number[] {
    return reorderByLevels(levels, start, end);
  }

  getMirror(codePoint: number): number | null {
    return mirrorGlyph(codePoint);
  }
}

export function createUax9Engine(): BidiEngine {
  return new Uax9Engine();
}
