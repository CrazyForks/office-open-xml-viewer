/**
 * IX2 public find-result shape, shared by all three viewers.
 *
 * `findText` returns an ordered list of {@link FindMatch}. Every match carries
 * its ordinal position (`matchIndex`, 0-based, document order — the same index
 * `findNext` / `findPrev` cycle through), the matched `text`, and a
 * format-specific `location`. The location is where the three formats
 * legitimately differ — a docx match lives on a page, a pptx match on a slide,
 * an xlsx match in a sheet cell — so `FindMatch` is generic over it rather than
 * forcing an artificial common shape. Each viewer instantiates it with its own
 * location type:
 *
 *   - `DocxViewer.findText` → `FindMatch<DocxMatchLocation>`  ({ page })
 *   - `PptxViewer.findText` → `FindMatch<PptxMatchLocation>`  ({ slide })
 *   - `XlsxViewer.findText` → `FindMatch<XlsxMatchLocation>`  ({ sheet, ref, … })
 *
 * The generic default is `unknown` so `FindMatch` can be referenced without a
 * type argument (e.g. in generic UI code) while each viewer's return type stays
 * precise.
 */
export interface FindMatch<Loc = unknown> {
  /** 0-based ordinal among all matches, in document order. This is the index
   *  `findNext`/`findPrev` make active, so a caller can correlate the array it
   *  got from `findText` with the active-match reported by navigation. */
  matchIndex: number;
  /** The text that matched (the query as it appears in the document — its
   *  original case, not the folded form used for case-insensitive matching). */
  text: string;
  /** Where the match is, in the format's own coordinates. */
  location: Loc;
}
