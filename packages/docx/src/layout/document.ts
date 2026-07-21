import type { BodyLayoutInput } from './body-layout-input.js';
import type { DocxDocumentModel } from '../types.js';
import { paginateBody } from './body-paginator.js';
import {
  attachDocumentLayoutVariants,
} from './document-layout-variants.js';
import { layoutVariantStoreOf } from './runtime-state.js';
import { normalizeLayoutOptions, type LayoutOptions } from './options.js';
import type { DocumentLayout, LayoutServices } from './types.js';

export function layoutDocumentInput(
  input: BodyLayoutInput,
  services: LayoutServices,
  options: LayoutOptions = normalizeLayoutOptions(undefined, Date.now()),
): DocumentLayout {
  return paginateBody(input, services, options);
}

export function ensureDocumentLayoutVariants(
  services: LayoutServices,
  defaultCurrentDateMs: number,
  resolveSource: () => Readonly<{
    model: DocxDocumentModel;
    input: BodyLayoutInput;
  }>,
): void {
  if (layoutVariantStoreOf(services)) return;
  const { model, input } = resolveSource();
  attachDocumentLayoutVariants({
    model,
    services,
    defaultCurrentDateMs,
    buildLayout: (options) => layoutDocumentInput(input, services, options),
  });
}
