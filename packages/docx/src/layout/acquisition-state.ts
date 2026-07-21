import type { ParagraphLayoutContext, StoryContext } from '../layout-context.js';
import {
  enterTableCellStoryContext,
  resolveParagraphLayoutContext,
} from '../layout-context.js';
import { getDefaultFontSize } from '../line-layout.js';
import type { DocParagraph } from '../types.js';
import type { BodyAcquisitionState, RetainedTableRecord } from './acquisition-context.js';
import { applyNumberingBodyOffset } from './numbering-marker.js';

export const BODY_STORY_CONTEXT: StoryContext = Object.freeze({
  story: 'body',
  containers: Object.freeze([]),
  lineNumberingEligible: true,
});

type ParagraphContextState = Pick<
  BodyAcquisitionState,
  'layoutSettings' | 'sectionLayout' | 'acquisitionInputs'
> & Partial<Pick<BodyAcquisitionState, 'layoutServices' | 'defaultTabPt'>>;

function applyNumberingContext(
  state: ParagraphContextState,
  paragraph: DocParagraph,
  context: ParagraphLayoutContext,
): ParagraphLayoutContext {
  return applyNumberingBodyOffset(context, {
    numbering: paragraph.numbering,
    ...(paragraph.numbering ? {
      markerInput: state.acquisitionInputs.numberingMarkerShapeInput(
        paragraph.numbering,
        getDefaultFontSize(paragraph),
      ),
    } : {}),
    authoredFirstIndentPt: paragraph.indentFirst,
    tabStops: paragraph.tabStops,
    defaultTabPt: state.defaultTabPt,
    service: state.layoutServices?.text,
  });
}

export function resolveBodyParagraphLayoutContext(
  state: ParagraphContextState,
  paragraph: DocParagraph,
): ParagraphLayoutContext {
  return applyNumberingContext(
    state,
    paragraph,
    resolveParagraphLayoutContext(
      state.layoutSettings,
      state.sectionLayout,
      BODY_STORY_CONTEXT,
      paragraph,
    ),
  );
}

export function resolveStateParagraphLayoutContext(
  state: ParagraphContextState & Partial<Pick<BodyAcquisitionState, 'storyContext'>>,
  paragraph: DocParagraph,
): ParagraphLayoutContext {
  return applyNumberingContext(
    state,
    paragraph,
    resolveParagraphLayoutContext(
      state.layoutSettings,
      state.sectionLayout,
      state.storyContext ?? BODY_STORY_CONTEXT,
      paragraph,
    ),
  );
}

export function withTableCellStory(state: BodyAcquisitionState): BodyAcquisitionState {
  return {
    ...state,
    storyContext: enterTableCellStoryContext(
      state.storyContext ?? BODY_STORY_CONTEXT,
    ),
  };
}

export function retainedTableRecord(
  state: BodyAcquisitionState,
  sourceIndex: number,
): RetainedTableRecord {
  const record = state.retainedTablesBySourceIndex?.get(sourceIndex);
  if (!record) throw new Error('Table placement requires retained table acquisition');
  return record;
}
