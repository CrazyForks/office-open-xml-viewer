import type { DocxDocumentModel } from './types.js';
import type { LayoutServices } from './layout/types.js';
import { attachDocumentLayoutVariants } from './layout/document-layout-variants.js';
import type { LayoutVariantStore } from './layout/variant-store.js';
import { layoutDocument } from './document-layout.js';

export interface RetainedRenderWorkerDocumentLayout {
  readonly model: DocxDocumentModel;
  readonly layoutServices: LayoutServices;
  readonly layoutVariants: LayoutVariantStore;
  readonly defaultCurrentDateMs: number;
}

/**
 * The worker keeps one document-scoped service graph and variant store. Keeping
 * their construction behind this pure seam lets parity tests execute the same
 * ownership wiring without importing the worker's WASM and `self` side effects.
 */
export function retainRenderWorkerDocumentLayout(
  model: DocxDocumentModel,
  layoutServices: LayoutServices,
  defaultCurrentDateMs: number,
): RetainedRenderWorkerDocumentLayout {
  const variants = attachDocumentLayoutVariants({
    model,
    services: layoutServices,
    defaultCurrentDateMs,
    buildLayout: (options) => layoutDocument(model, layoutServices, options),
  });
  return Object.freeze({
    model,
    layoutServices,
    layoutVariants: variants.store,
    defaultCurrentDateMs,
  });
}
