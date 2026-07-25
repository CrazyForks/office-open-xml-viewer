import type { DeepReadonly, DocumentLayout } from './layout/types.js';

/** Project retained bookmark ownership to its first physical page. */
export function buildBookmarkPageMap(
  layout: DocumentLayout | DeepReadonly<DocumentLayout>,
): Map<string, number> {
  const map = new Map<string, number>();
  for (const page of layout.pages) {
    for (const bookmark of page.bookmarkStarts) {
      if (bookmark.name !== '' && !map.has(bookmark.name)) {
        map.set(bookmark.name, page.pageIndex);
      }
    }
  }
  return map;
}

export function resolveBookmarkPage(
  map: Map<string, number>,
  bookmarkName: string,
): number | undefined {
  return map.get(bookmarkName);
}
