/**
 * CSS colours used by in-document search overlays.
 *
 * The values are applied as the overlay backgrounds verbatim. Use an alpha
 * colour (`rgba(...)`, 8-digit hex, or `color-mix(...)`) when the rendered text
 * should remain visible through the highlight.
 */
export interface FindHighlightColors {
  /** Background for every match except the active one. */
  match?: string;
  /** Background for the match selected by findNext/findPrev. */
  active?: string;
}
