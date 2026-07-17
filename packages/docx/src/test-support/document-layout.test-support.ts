import type { DocumentLayout } from '../layout/types.js';
import type { BodyElement, DocxDocumentModel, SectionProps } from '../types.js';
import { createLayoutServices, layoutDocument } from '../renderer.js';

const EMPTY_STORY = Object.freeze({ default: null, first: null, even: null });

export function layoutBodyModel(
  body: readonly BodyElement[],
  section: SectionProps,
  measureContext: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  fontFamilyClasses: Readonly<Record<string, string>> = {},
): DocumentLayout {
  const model = {
    body: [...body],
    section,
    headers: EMPTY_STORY,
    footers: EMPTY_STORY,
    footnotes: [],
    endnotes: [],
    fontFamilyClasses: { ...fontFamilyClasses },
  } as unknown as DocxDocumentModel;
  const services = createLayoutServices(model, { measureContext });
  return layoutDocument(model, services, { currentDateMs: 0 });
}
