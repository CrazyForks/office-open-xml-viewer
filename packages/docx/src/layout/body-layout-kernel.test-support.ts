import type {
  BodyLayoutKernel,
  BodyLayoutSession,
  BodyLayoutSessionInput,
  NoteLayoutAcquisitionInput,
  StoryLayoutAcquisitionInput,
} from './body-layout-kernel.js';
import { attachBodyLayoutKernel } from './runtime-state.js';
import type { LayoutOptions } from './options.js';
import type {
  LayoutServices,
  NoteLayout,
  StoryLayout,
} from './types.js';

type SyntheticBodyLayoutSession = Omit<BodyLayoutSession, 'layoutStory' | 'layoutNotes'> & Readonly<{
  layoutStory?: BodyLayoutSession['layoutStory'];
  layoutNotes?: BodyLayoutSession['layoutNotes'];
  measureStoryExtent?: (request: Readonly<{
    source: StoryLayoutAcquisitionInput['source'];
    pageIndex: number;
    section: StoryLayoutAcquisitionInput['section'];
    availableInlineExtentPt: number;
  }>) => number;
  measureFootnoteReserve?: (request: Readonly<{
    referenceIds: readonly string[];
    availableInlineExtentPt: number;
    firstOnPage: boolean;
  }>) => number;
}>;

export interface SyntheticBodyLayoutKernel {
  openBodyLayoutSession(
    input: BodyLayoutSessionInput,
    services: LayoutServices,
    options: LayoutOptions,
  ): SyntheticBodyLayoutSession;
}

function emptyStory(
  request: StoryLayoutAcquisitionInput,
  advancePt: number,
): StoryLayout {
  const bounds = Object.freeze({
    xPt: request.container.bounds.xPt,
    yPt: request.container.bounds.yPt,
    widthPt: request.container.bounds.widthPt,
    heightPt: advancePt,
  });
  return Object.freeze({
    story: request.source.story,
    flowBounds: bounds,
    inkBounds: bounds,
    blocks: Object.freeze([]),
    advancePt,
  });
}

function syntheticNotes(
  request: NoteLayoutAcquisitionInput,
  advancePt: number,
): readonly NoteLayout[] {
  if (advancePt <= 0 || request.referenceIds.length === 0) return Object.freeze([]);
  const source = Object.freeze({
    story: request.kind,
    storyInstance: request.referenceIds[0]!,
    path: Object.freeze([]),
  });
  const bounds = Object.freeze({
    xPt: request.container.bounds.xPt,
    yPt: request.container.bounds.yPt,
    widthPt: request.container.bounds.widthPt,
    heightPt: advancePt,
  });
  const story = Object.freeze({
    story: request.kind,
    flowBounds: bounds,
    inkBounds: bounds,
    blocks: Object.freeze([]),
    advancePt,
  });
  return Object.freeze([Object.freeze({
    kind: 'note',
    id: `${request.kind}:${encodeURIComponent(request.referenceIds.join(','))}:page:${request.pageIndex}`,
    source,
    flowDomainId: request.container.id,
    ordinaryFlow: true,
    flowBounds: bounds,
    inkBounds: bounds,
    clipBounds: request.container.bounds,
    advancePt,
    separator: Object.freeze([]),
    story,
  })]);
}

/**
 * Adapt focused paginator stubs to the retained B1 story/note contract. Numeric
 * reserve callbacks remain test-only inputs; production never sees this seam.
 */
export function attachSyntheticBodyLayoutKernel(
  services: LayoutServices,
  kernel: SyntheticBodyLayoutKernel,
): void {
  const productionKernel: BodyLayoutKernel = {
    openBodyLayoutSession(input, runtimeServices, options) {
      const synthetic = kernel.openBodyLayoutSession(input, runtimeServices, options);
      const {
        measureStoryExtent,
        measureFootnoteReserve,
        ...session
      } = synthetic;
      return Object.freeze({
        ...session,
        layoutStory: synthetic.layoutStory ?? ((request) => emptyStory(
          request,
          measureStoryExtent?.({
            source: request.source,
            pageIndex: request.pageIndex,
            section: request.section,
            availableInlineExtentPt: request.container.bounds.widthPt,
          }) ?? 0,
        )),
        layoutNotes: synthetic.layoutNotes ?? ((request) => syntheticNotes(
          request,
          measureFootnoteReserve?.({
            referenceIds: request.referenceIds,
            availableInlineExtentPt: request.container.bounds.widthPt,
            firstOnPage: request.firstOnPage,
          }) ?? 0,
        )),
      });
    },
  };
  attachBodyLayoutKernel(services, productionKernel);
}
