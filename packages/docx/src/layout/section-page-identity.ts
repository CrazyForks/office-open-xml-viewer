export interface SectionContentPage {
  readonly pageIndex: number;
  readonly sectionRegions: readonly Readonly<{
    sectionOccurrenceId: string;
  }>[];
}

export interface SectionOwnedPage {
  readonly sectionOccurrenceId: string;
}

/** The first physical page containing body content from each retained section.
 * Content/numbering arithmetic may use this even when another section owns the page. */
export function sectionContentFirstAppearancePageIndices(
  pages: readonly SectionContentPage[],
): ReadonlyMap<string, number> {
  const firstAppearance = new Map<string, number>();
  for (const page of pages) {
    for (const region of page.sectionRegions) {
      if (!firstAppearance.has(region.sectionOccurrenceId)) {
        firstAppearance.set(region.sectionOccurrenceId, page.pageIndex);
      }
    }
  }
  return firstAppearance;
}

/** Whether this is the first physical page whose page-level owner is its section. */
export function isFirstSectionOwnedPage(
  pages: readonly SectionOwnedPage[],
  pageIndex: number,
): boolean {
  const page = pages[pageIndex];
  if (!page) return false;
  return pageIndex === 0
    || pages[pageIndex - 1]?.sectionOccurrenceId !== page.sectionOccurrenceId;
}
