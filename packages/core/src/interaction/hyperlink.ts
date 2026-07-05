/**
 * Shared hyperlink model + URL sanitisation for docx / pptx / xlsx (IX1).
 *
 * All three formats carry the same two ECMA-376 concepts:
 *   - an **external** hyperlink — an absolute URL resolved from a relationship
 *     part target (`document.xml.rels` for docx §17.16.22, the slide rels for
 *     pptx §21.1.2.3.5, the worksheet rels for xlsx §18.3.1.47), with
 *     `TargetMode="External"`.
 *   - an **internal** hyperlink — a jump within the document itself:
 *     docx `w:anchor` -> a `<w:bookmarkStart w:name>` (§17.16.23), pptx
 *     `action="ppaction://hlinksldjump"` -> a slide, xlsx `location` -> a defined
 *     name or a `Sheet!A1` cell reference.
 *
 * The parsers (Rust, one per format) do the format-specific rels lookup and hand
 * each run / shape / cell a {@link HyperlinkTarget}. Everything downstream — the
 * text-layer overlay, the viewer default click behaviour, and any integrator
 * callback — is format-agnostic and consumes this one shape. Keeping the type +
 * the pure `sanitizeHyperlinkUrl` predicate here (not duplicated per package)
 * follows the cross-package unification principle: a scheme-allowlist bug fixed
 * once is fixed everywhere.
 */

/**
 * A resolved hyperlink attached to a run, shape, or cell.
 *
 *   - `external` — `url` is the raw target as authored in the file. It is NOT
 *     guaranteed safe; run it through {@link sanitizeHyperlinkUrl} before
 *     navigating. It is kept verbatim here so an integrator can apply its own
 *     policy (e.g. allow `file:` on a trusted intranet viewer).
 *   - `internal` — `ref` is the in-document destination, verbatim from the file:
 *       docx: the bookmark name (`w:anchor`).
 *       pptx: the internal action (e.g. `ppaction://hlinksldjump`), with the
 *             resolved 0-based `slideIndex` when the rels target names a slide.
 *       xlsx: the `location` string (a defined name or `Sheet1!A1`).
 */
export type HyperlinkTarget =
  | { kind: 'external'; url: string }
  | { kind: 'internal'; ref: string; slideIndex?: number };

/**
 * URL schemes permitted for external hyperlink navigation by default.
 *
 * `http` / `https` cover the web; `mailto` and `tel` are inert launch handlers
 * with no script-execution surface. Everything else — notably `javascript:`,
 * `data:`, `vbscript:`, `file:`, `about:`, `blob:` — is refused, because a
 * malicious document could otherwise smuggle script execution or local-file
 * disclosure through a link the user is invited to click. This mirrors the
 * conservative allowlist browsers apply to `target=_blank` navigations and is
 * consistent with the library's typed-error / fail-safe stance (RB series):
 * when in doubt, do nothing rather than execute attacker-controlled input.
 */
export const DEFAULT_ALLOWED_HYPERLINK_SCHEMES: readonly string[] = [
  'http',
  'https',
  'mailto',
  'tel',
];

/**
 * Extract the lowercased URL scheme (the token before the first `:`), or `null`
 * when the string carries no scheme. A scheme is `ALPHA *( ALPHA / DIGIT / "+"
 * / "-" / "." )` per RFC 3986 §3.1. A leading-colon or empty string has no
 * scheme.
 *
 * All ASCII whitespace and C0 control characters (code points 0x00–0x20),
 * including the embedded TAB / LF / CR that browsers strip from URLs — the trick
 * attackers use to hide `java<TAB>script:` from a naive prefix check — are
 * removed before scanning, so `"  java<LF>script:alert(1)"` is correctly seen as
 * the `javascript` scheme and rejected.
 */
export function hyperlinkUrlScheme(url: string): string | null {
  // Drop every ASCII control char and space (code points 0x00–0x20). Browsers
  // ignore embedded TAB/LF/CR and leading spaces when resolving a URL, so a
  // checker that leaves them in would under-report the real scheme. A code-point
  // filter is used (rather than a control-char regex literal) to keep this
  // source file free of embedded control characters.
  let cleaned = '';
  for (const ch of url) {
    const cp = ch.codePointAt(0);
    if (cp !== undefined && cp > 0x20) cleaned += ch;
  }
  const m = /^([a-zA-Z][a-zA-Z0-9+.-]*):/.exec(cleaned);
  return m ? m[1].toLowerCase() : null;
}

/**
 * Validate an external hyperlink URL against a scheme allowlist and return the
 * URL to navigate to, or `null` if it must be blocked.
 *
 * A URL with **no scheme** (a relative or protocol-relative reference such as
 * `"page.html"` or `"//example.com/x"`) is treated as safe and returned as-is:
 * it cannot name a dangerous scheme, and it resolves against the current
 * document origin exactly like any relative link. A URL whose scheme is not in
 * `allowed` (default {@link DEFAULT_ALLOWED_HYPERLINK_SCHEMES}) returns `null`
 * so callers can no-op the click.
 *
 * This is a pure predicate: it neither navigates nor mutates. Comparison is
 * case-insensitive on the scheme; the returned string is the original,
 * untrimmed input (so query/fragment/casing are preserved for navigation).
 *
 * @param url     the raw external target from the file.
 * @param allowed lowercase scheme allowlist; defaults to the safe set.
 * @returns the URL to open, or `null` when the scheme is disallowed.
 */
export function sanitizeHyperlinkUrl(
  url: string,
  allowed: readonly string[] = DEFAULT_ALLOWED_HYPERLINK_SCHEMES,
): string | null {
  if (url === '') return null;
  const scheme = hyperlinkUrlScheme(url);
  // No scheme -> relative/anchor reference, inherently same-origin-safe.
  if (scheme === null) return url;
  return allowed.includes(scheme) ? url : null;
}

/**
 * The default action a viewer takes for an **external** hyperlink click when
 * the integrator supplies no `onHyperlinkClick` handler: sanitise the URL and,
 * if allowed, open it in a new tab with `noopener,noreferrer` so the opened page
 * gets no `window.opener` handle back into this document. A blocked scheme is a
 * silent no-op (returns `false`) — the click does nothing rather than navigate
 * somewhere dangerous.
 *
 * Internal targets are intentionally NOT handled here: the in-document jump
 * (page / slide / cell) is format-specific and lives in each viewer.
 *
 * Split out (not inlined in three viewers) so the "open in new tab, drop opener,
 * refuse unsafe schemes" policy is defined once. `win` is injected for tests;
 * defaults to the ambient `window`.
 *
 * @returns `true` if navigation was initiated, `false` if the URL was blocked.
 */
export function openExternalHyperlink(
  url: string,
  allowed: readonly string[] = DEFAULT_ALLOWED_HYPERLINK_SCHEMES,
  win: Pick<Window, 'open'> | undefined = typeof window !== 'undefined' ? window : undefined,
): boolean {
  const safe = sanitizeHyperlinkUrl(url, allowed);
  if (safe === null || !win) return false;
  win.open(safe, '_blank', 'noopener,noreferrer');
  return true;
}
