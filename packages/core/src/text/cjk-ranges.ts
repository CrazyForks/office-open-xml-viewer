// Shared classifier for CJK / ideographic code points that permit a line break
// (and a justification stretch) between adjacent characters, without any
// intervening whitespace. This is the single source of truth consumed by every
// renderer; previously the same four ranges were hand-duplicated in
// packages/{pptx,docx,xlsx} and had drifted (the Hangul upper bound was D7FF in
// one copy and D7AF in two others — both wrong, see below).
//
// ── Canonical ranges ────────────────────────────────────────────────────────
//   U+3000–U+9FFF  Ideographic space, CJK Symbols & Punctuation, Hiragana,
//                  Katakana, CJK Unified Ideographs (incl. Extension A and the
//                  Kangxi/CJK symbol blocks that fall inside this span).
//   U+AC00–U+D7A3  Hangul Syllables.
//   U+F900–U+FAFF  CJK Compatibility Ideographs.
//   U+FF00–U+FFEF  Halfwidth and Fullwidth Forms.
//
// ── Why U+D7A3 (not D7AF / D7FF) is the Hangul upper bound ───────────────────
// The Unicode "Hangul Syllables" block is exactly U+AC00–U+D7A3 (the 11,172
// precomposed modern syllable blocks). The code points above it are NOT
// standalone CJK break units and must be excluded:
//   • U+D7A4–U+D7AF  unassigned (no characters).
//   • U+D7B0–U+D7FF  "Hangul Jamo Extended-B" — combining/conjoining jamo, i.e.
//                    pieces that attach to a base syllable. A combining mark is
//                    never an independent line-break or justification unit, so
//                    it must stay with its base and is excluded here too.
// The earlier copies' D7AF (included 12 unassigned code points) and D7FF
// (additionally swept in all of Jamo Extended-B) were both over-broad; D7A3 is
// the strict, correct boundary. Behaviour is unchanged for real content: no
// sample exercises U+D7A4–U+D7FF.
//
// ── U+3000 (ideographic space): wrap vs. justify ────────────────────────────
// U+3000 IS intentionally part of this range. For WRAP/BREAK callers that is
// required: a line may break around an ideographic space like any other CJK
// glyph. For JUSTIFY callers it is harmless: a justifier classifies U+3000 as
// whitespace (it matches /\s/) and stretches it as an inter-word gap BEFORE it
// would ever reach this predicate, so including U+3000 here never causes a
// U+3000 unit to be counted a second time as an inter-CJK gap. Renderers can
// therefore share this one predicate for both purposes. (Historically the pptx
// justify copy started its range at U+3001 to "exclude" U+3000; that exclusion
// was redundant for the reason just given, and is dropped here.)

/**
 * True when `cp` is a CJK / ideographic code point that allows a line break —
 * and a justification stretch — at the boundary with an adjacent character,
 * with no intervening whitespace required.
 *
 * @param cp A Unicode scalar value (e.g. from `String.prototype.codePointAt`).
 */
export function isCjkBreakChar(cp: number): boolean {
  return (
    (cp >= 0x3000 && cp <= 0x9fff) || // CJK punctuation/symbols, kana, Unified (incl. Ext-A)
    (cp >= 0xac00 && cp <= 0xd7a3) || // Hangul Syllables
    (cp >= 0xf900 && cp <= 0xfaff) || // CJK Compatibility Ideographs
    (cp >= 0xff00 && cp <= 0xffef) // Halfwidth/Fullwidth Forms
  );
}
