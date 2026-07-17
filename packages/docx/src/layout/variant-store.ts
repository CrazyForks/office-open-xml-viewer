import { deepFreezeDocumentLayout } from './invariants.js';
import { layoutOptionsKey, type LayoutOptions } from './options.js';
import type { DeepReadonly, DocumentLayout, LayoutPage, LayoutServices } from './types.js';

export type DocumentLayoutBuilder = (
  options: LayoutOptions,
) => DocumentLayout | DeepReadonly<DocumentLayout>;

export interface DocumentLayoutSelection {
  readonly key: string;
  readonly options: LayoutOptions;
  readonly layout: DeepReadonly<DocumentLayout>;
}

export function requireLayoutPage(
  layout: DeepReadonly<DocumentLayout>,
  pageIndex: number,
): DeepReadonly<LayoutPage> {
  if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= layout.pages.length) {
    throw new RangeError(`Page index ${pageIndex} out of range (count: ${layout.pages.length})`);
  }
  return layout.pages[pageIndex] as DeepReadonly<LayoutPage>;
}

/**
 * Document-scoped layout cache. The key deliberately excludes paint-only facts
 * such as scale, DPR, and color: only acquisition inputs may select geometry.
 */
export class LayoutVariantStore {
  readonly #services: LayoutServices;
  readonly #build: DocumentLayoutBuilder;
  readonly #variants = new Map<string, DeepReadonly<DocumentLayout>>();
  readonly #defaultOptions: LayoutOptions;
  readonly #defaultKey: string;

  constructor(
    services: LayoutServices,
    defaultOptions: LayoutOptions,
    build: DocumentLayoutBuilder,
  ) {
    this.#services = services;
    this.#defaultOptions = Object.freeze({ ...defaultOptions });
    this.#defaultKey = layoutOptionsKey(this.#defaultOptions, this.#services);
    this.#build = build;
  }

  get defaultLayout(): DeepReadonly<DocumentLayout> {
    return this.layoutFor(this.#defaultOptions);
  }

  layoutFor(options: LayoutOptions): DeepReadonly<DocumentLayout> {
    return this.select(options).layout;
  }

  select(options: LayoutOptions): DocumentLayoutSelection {
    const normalized = Object.isFrozen(options)
      ? options
      : Object.freeze({ ...options });
    const key = layoutOptionsKey(normalized, this.#services);
    let layout = this.#variants.get(key);
    if (!layout) {
      layout = deepFreezeDocumentLayout(this.#build(normalized) as DocumentLayout);
      this.#variants.set(key, layout);
    }
    return Object.freeze({ key, options: normalized, layout });
  }

  selectPage(
    options: LayoutOptions,
    pageIndex: number,
  ): Readonly<{
    key: string;
    options: LayoutOptions;
    layout: DeepReadonly<DocumentLayout>;
    page: DeepReadonly<LayoutPage>;
  }> {
    const selection = this.select(options);
    return Object.freeze({
      ...selection,
      page: requireLayoutPage(selection.layout, pageIndex),
    });
  }

  isDefault(options: LayoutOptions): boolean {
    return layoutOptionsKey(options, this.#services) === this.#defaultKey;
  }
}
