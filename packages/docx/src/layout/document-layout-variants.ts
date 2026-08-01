import { layoutParseErrorPage } from './error-page.js';
import type { LayoutSourceStore } from './layout-source-store.js';
import {
  layoutOptionsForRender,
  type LayoutOptions,
  type LayoutRenderSelectionInput,
} from './options.js';
import {
  attachLayoutVariantStore,
  layoutVariantStoreOf,
} from './runtime-state.js';
import type { DeepReadonly, DocumentLayout, LayoutPage, LayoutServices } from './types.js';
import { LayoutVariantStore, type DocumentLayoutBuilder } from './variant-store.js';

export type DocumentLayoutVariantFactoryInput = Readonly<{
  readonly source: LayoutSourceStore;
  readonly services: LayoutServices;
  readonly defaultCurrentDateMs: number;
  readonly buildLayout: DocumentLayoutBuilder;
}>;

export interface AttachedDocumentLayoutVariants {
  readonly store: LayoutVariantStore;
  readonly defaultOptions: LayoutOptions;
}

export function attachDocumentLayoutVariants(
  input: DocumentLayoutVariantFactoryInput,
): AttachedDocumentLayoutVariants {
  const { services, defaultCurrentDateMs, buildLayout } = input;
  const defaultOptions = layoutOptionsForRender({ defaultCurrentDateMs });
  const fatalParse = input.source.fatalParse;
  const parseError = fatalParse === null
    ? null
    : layoutParseErrorPage(
        fatalParse.message,
        fatalParse.pageSize,
        services.text,
      );
  const store = new LayoutVariantStore(
    services,
    defaultOptions,
    parseError === null ? buildLayout : () => parseError,
  );
  attachLayoutVariantStore(services, store);
  return Object.freeze({ store, defaultOptions });
}

export function selectDocumentLayoutPage(
  services: LayoutServices,
  input: LayoutRenderSelectionInput,
  pageIndex: number,
): Readonly<{
  key: string;
  options: LayoutOptions;
  layout: DeepReadonly<DocumentLayout>;
  page: DeepReadonly<LayoutPage>;
}> {
  const store = layoutVariantStoreOf(services);
  if (!store) throw new Error('Document layout variant store is not attached to the supplied services');
  return store.selectPage(layoutOptionsForRender(input), pageIndex);
}
