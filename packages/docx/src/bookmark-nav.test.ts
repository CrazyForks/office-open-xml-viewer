import { describe, it, expect } from 'vitest';
import type { DocumentLayout } from './layout/types.js';
import { buildBookmarkPageMap, resolveBookmarkPage } from './bookmark-nav';

/** Minimal paginated paragraph carrying the given bookmark names. Only the
 *  fields the map builder reads (`type`, `bookmarks`) matter; the rest are
 *  filled to satisfy the type without affecting resolution. */
function layout(bookmarksByPage: readonly (readonly string[])[]): DocumentLayout {
  return {
    pages: bookmarksByPage.map((names, pageIndex) => ({
      pageIndex,
      bookmarkStarts: names.map((name) => ({ name })),
    })),
    diagnostics: [],
  } as unknown as DocumentLayout;
}

describe('buildBookmarkPageMap', () => {
  it('maps a bookmark to the 0-based index of the page carrying it', () => {
    const pages = layout([['_Toc_intro'], ['_Toc_methods'], ['_Toc_results']]);
    const map = buildBookmarkPageMap(pages);
    expect(map.get('_Toc_intro')).toBe(0);
    expect(map.get('_Toc_methods')).toBe(1);
    expect(map.get('_Toc_results')).toBe(2);
  });

  it('maps multiple bookmarks that share a paragraph to the same page', () => {
    const pages = layout([[], ['a', 'b', 'c']]);
    const map = buildBookmarkPageMap(pages);
    expect(map.get('a')).toBe(1);
    expect(map.get('b')).toBe(1);
    expect(map.get('c')).toBe(1);
  });

  it('resolves a name repeated across pages to the FIRST (earliest) page', () => {
    // A paragraph split across a page break carries its bookmark on both slices;
    // the destination is where the paragraph begins.
    const pages = layout([['dup'], ['dup']]);
    expect(buildBookmarkPageMap(pages).get('dup')).toBe(0);
  });

  it('returns undefined for a name that appears in no paragraph', () => {
    const map = buildBookmarkPageMap(layout([['known']]));
    expect(map.get('missing')).toBeUndefined();
  });

  it('honors bookmarks nested in a table cell paragraph', () => {
    const pages = layout([[], ['cellmark']]);
    expect(buildBookmarkPageMap(pages).get('cellmark')).toBe(1);
  });

  it('ignores empty-string bookmark names', () => {
    const map = buildBookmarkPageMap(layout([['']]));
    expect(map.size).toBe(0);
  });
});

describe('resolveBookmarkPage', () => {
  it('is a thin lookup over the built map', () => {
    const map = buildBookmarkPageMap(layout([[], ['x']]));
    expect(resolveBookmarkPage(map, 'x')).toBe(1);
    expect(resolveBookmarkPage(map, 'nope')).toBeUndefined();
  });
});

// M2: main mode builds the bookmark map from the paginated pages; worker mode
// SERIALIZES it into `DocumentMeta.bookmarkPages` (a `[name, page][]` array, the
// Map→array form that survives `postMessage`) and the main-thread proxy
// RECONSTRUCTS it with `new Map(meta.bookmarkPages)`. A drop or re-order in that
// round-trip would make an internal `<w:anchor>` land on the wrong page in worker
// mode only. Pin that the serialized-then-reconstructed map equals the directly
// built map, entry-for-entry.
describe('bookmark map — main/worker serialization equivalence (M2)', () => {
  const pages = layout([
    ['_Toc_intro'],
    ['_Toc_methods', 'alias'],
    ['cellmark'],
  ]);

  it('round-trips the map through the worker-meta [name, page][] wire form unchanged', () => {
    const mainMap = buildBookmarkPageMap(pages); // main mode
    // Worker: render-worker.ts does `bookmarkPages: [...buildBookmarkPageMap(pages)]`.
    const wireForm: [string, number][] = [...buildBookmarkPageMap(pages)];
    // Main proxy: document.ts does `new Map(this._meta.bookmarkPages)`.
    const workerMap = new Map(wireForm);

    // Same keys, same values, same size — a faithful round-trip.
    expect(workerMap.size).toBe(mainMap.size);
    expect(mainMap.size).toBeGreaterThan(0); // non-degenerate
    for (const [name, page] of mainMap) {
      expect(workerMap.get(name)).toBe(page);
    }
    // And every resolved anchor agrees across the two modes.
    for (const name of ['_Toc_intro', '_Toc_methods', 'alias', 'cellmark', 'missing']) {
      expect(resolveBookmarkPage(workerMap, name)).toBe(resolveBookmarkPage(mainMap, name));
    }
  });
});
