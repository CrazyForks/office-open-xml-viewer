import { defineCompatibilityRule } from './compatibility.js';

export const WORD_NEUTRAL_SCRIPT_ATTACHMENT = defineCompatibilityRule({
  id: 'word-neutral-script-attachment',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/compatibility.test.ts#keeps neutral characters attached to the active script slice',
  },
  description: 'Weak and neutral non-letter characters stay with the active complex-script slice instead of opening additional formatting segments.',
});

export const WORD_RTL_RUN_AMBIGUOUS_CLASS_OVERRIDE = defineCompatibilityRule({
  id: 'word-rtl-run-ambiguous-class-override',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/bidi-line.test.ts#keeps LTR word order for English text in rtl-marked runs',
  },
  description: 'Model an rtl-marked run as a higher-level UAX #9 override for punctuation and symbols only, leaving whitespace and strong letters at their ordinary classes.',
});

export const WORD_KASHIDA_FINAL_FORM_PRIORITY = defineCompatibilityRule({
  id: 'word-kashida-final-form-priority',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/kashida-priority.test.ts#uses the BaRa join (Beh->Yeh) over the final-letter join in بين',
  },
  description: 'Apply the measured kashida final-letter priority classes only at a word-final following letter instead of copying the broader Qt final-form conditions.',
});

export const WORD_VERTICAL_TU_CORNER_PLACEMENT = defineCompatibilityRule({
  id: 'word-vertical-tu-corner-placement',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/vertical-text.test.ts#does NOT ink-centre a substituted Tu comma even when ink metrics are present',
  },
  description: 'Keep a substituted vertical Tu comma or full stop at the font-designed upper-right cell position rather than ink-centering it geometrically.',
});

const RTL_AMBIGUOUS_CHARACTER = /[\p{P}\p{S}]/u;

export function wordRtlAmbiguousCharacter(character: string): boolean {
  return RTL_AMBIGUOUS_CHARACTER.test(character);
}

export function wordNeutralAttachesToActiveScript(character: string): boolean {
  return !/\p{L}/u.test(character);
}

export function wordKashidaFinalFormApplies(
  beforeCodePointIndex: number,
  lastLetterCodePointIndex: number,
): boolean {
  return beforeCodePointIndex === lastLetterCodePointIndex;
}

export function wordPreservesVerticalTuCorner(
  substitutedPunctuationCodePoint: number | null,
): boolean {
  return substitutedPunctuationCodePoint !== null;
}
