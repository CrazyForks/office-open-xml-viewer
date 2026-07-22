/**
 * Same-font Canvas-vs-Word horizontal advance bias (Regime A).
 *
 * This module corrects the measurement gap between Canvas `measureText` and
 * Word's own layout advances for the SAME resolved face. It deliberately does
 * NOT emulate the metrics of a font the host does not have: when a requested
 * face is absent, the viewer substitutes the host's fallback face and reflows —
 * exactly like Word does when it opens a document whose fonts are missing.
 * Line breaks and page counts may then differ from the authoring machine; that
 * is accepted behavior, not a defect (product decision 2026-07-12; issue #855
 * closed as won't-fix). Cross-font per-script advance profiles that faked the
 * requested face's metrics on top of a substitute (PR #979 "Regime B") were
 * reverted because per-family metric tables cannot scale across the open set
 * of document fonts.
 *
 * The profile keys the selected Canvas family route. A family belongs here only
 * when its bias has documented Word/PDF provenance; unknown families remain a
 * strict zero no-op so existing layout is byte-stable.
 */

interface FontBiasProfile {
  readonly test: (family: string) => boolean;
  readonly biasEm: number;
}

/**
 * Canvas-vs-Word horizontal advance bias in em per glyph. This is a line-fit
 * allowance for the gap between Canvas `measureText` advances and Word's own
 * layout advances for the selected family route, not a glyph transform. The
 * allowance is backend-agnostic by design. The Georgia profile is applied only
 * to segments routed through Georgia; per-character OOXML font-slot routing
 * splits East Asian punctuation into its actual family before this lookup.
 * The value remains the public Word/PDF-calibrated Chromium allowance from the
 * pre-refactor renderer.
 */
const FONT_BIAS_PROFILES: ReadonlyArray<FontBiasProfile> = [
  {
    test: (family) => family === 'georgia',
    biasEm: 0.0105,
  },
];

function normalizeFamily(family: string | null | undefined): string {
  return (family ?? '')
    .trim()
    .replace(/^["']|["']$/g, '')
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

/** Expected Canvas-over-Word advance bias, in em per glyph. */
export function fontAdvanceBiasEm(family: string | null | undefined): number {
  const normalized = normalizeFamily(family);
  for (const profile of FONT_BIAS_PROFILES) {
    if (profile.test(normalized)) return profile.biasEm;
  }
  return 0;
}
