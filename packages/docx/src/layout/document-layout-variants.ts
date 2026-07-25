import type { DocxDocumentModel } from '../types.js';
import { layoutParseErrorPage } from './error-page.js';
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

export interface DocumentLayoutVariantFactoryInput {
  readonly model: DocxDocumentModel;
  readonly services: LayoutServices;
  readonly defaultCurrentDateMs: number;
  readonly buildLayout: DocumentLayoutBuilder;
}

export interface AttachedDocumentLayoutVariants {
  readonly store: LayoutVariantStore;
  readonly defaultOptions: LayoutOptions;
}

export function attachDocumentLayoutVariants({
  model,
  services,
  defaultCurrentDateMs,
  buildLayout,
}: DocumentLayoutVariantFactoryInput): AttachedDocumentLayoutVariants {
  const defaultOptions = layoutOptionsForRender({ defaultCurrentDateMs });
  const parseError = model.parseError === undefined
    ? null
    : layoutParseErrorPage(
        model.parseError,
        { widthPt: model.section.pageWidth, heightPt: model.section.pageHeight },
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
