import type {
  DrawingMLCollisionEntryPt,
  ParagraphLayout,
  WrapExclusion,
} from './types.js';

export interface ParagraphWrapRegistry {
  readonly flowDomainId: string;
  readonly collisions: readonly DrawingMLCollisionEntryPt[];
  readonly exclusions: readonly WrapExclusion[];
}

const registriesByOwner = new WeakMap<object, Map<string, ParagraphWrapRegistry>>();

export const TRANSIENT_TABLE_FINAL_FRAME_EXCLUSION_PREFIX = 'table-final-frame:';

/** Reacquisition replaces final-frame table probes but preserves every other
 * inherited authority. Ownership comes from retained drawings so parser-owned
 * and public compatibility anchors follow the same occurrence contract. */
export function inheritedParagraphAuthorityForReacquisition(
  layout: ParagraphLayout,
): Readonly<{
  exclusions: readonly WrapExclusion[];
  collisions: readonly DrawingMLCollisionEntryPt[];
}> {
  const ownedOccurrences = new Set(layout.drawings.flatMap((drawing) => {
    const occurrenceId = drawing.anchorLayer?.acquisitionOccurrenceId
      ?? drawing.anchorLayer?.occurrenceId;
    return occurrenceId === undefined ? [] : [occurrenceId];
  }));
  return Object.freeze({
    exclusions: Object.freeze(layout.exclusions.filter((exclusion) =>
      !exclusion.id.startsWith(TRANSIENT_TABLE_FINAL_FRAME_EXCLUSION_PREFIX)
      && (
        exclusion.anchorOccurrenceId === undefined
        || !ownedOccurrences.has(exclusion.anchorOccurrenceId)
      ))),
    collisions: Object.freeze((layout.anchorCollisions ?? []).filter((entry) =>
      !ownedOccurrences.has(entry.occurrenceId))),
  });
}

function ownedParagraphOccurrenceIds(layout: ParagraphLayout): ReadonlySet<string> {
  return new Set((layout.anchorFrames ?? []).flatMap((frame) =>
    frame.status === 'resolved' ? [frame.occurrenceId] : []));
}

export function ownedParagraphAnchorCollisions(
  layout: ParagraphLayout,
): readonly DrawingMLCollisionEntryPt[] {
  const ownedOccurrences = ownedParagraphOccurrenceIds(layout);
  const collisions = (layout.anchorCollisions ?? []).filter((entry) =>
    ownedOccurrences.has(entry.occurrenceId));
  const retainedOccurrences = new Set(collisions.map((entry) => entry.occurrenceId));
  for (const occurrenceId of ownedOccurrences) {
    if (!retainedOccurrences.has(occurrenceId)) {
      throw new Error(`Paragraph anchor omitted collision geometry: ${occurrenceId}`);
    }
  }
  return Object.freeze(collisions);
}

/** Effective exclusions inherited from earlier paragraphs remain observable on
 * the layout but must not be committed a second time. */
export function ownedParagraphWrapExclusions(
  layout: ParagraphLayout,
): readonly WrapExclusion[] {
  const ownedOccurrences = ownedParagraphOccurrenceIds(layout);
  return Object.freeze(layout.exclusions.filter((exclusion) =>
    exclusion.anchorOccurrenceId !== undefined
      && ownedOccurrences.has(exclusion.anchorOccurrenceId)));
}

/** Flow-local occurrence registry shared by sequential paragraph acquisitions.
 * Object collisions and text exclusions remain distinct retained authorities. */
export function createParagraphWrapRegistry(
  flowDomainId: string,
): ParagraphWrapRegistry {
  return Object.freeze({
    flowDomainId,
    collisions: Object.freeze([]),
    exclusions: Object.freeze([]),
  });
}

/** Owner identity scopes sequential paragraph authority to one acquisition
 * state. The registry itself stays immutable so rejected candidates cannot
 * mutate authority observed by a later paragraph. */
export function accessParagraphWrapRegistry(
  owner: object,
  flowDomainId: string,
): ParagraphWrapRegistry {
  let byFlowDomain = registriesByOwner.get(owner);
  if (!byFlowDomain) {
    byFlowDomain = new Map();
    registriesByOwner.set(owner, byFlowDomain);
  }
  const existing = byFlowDomain.get(flowDomainId);
  if (existing) return existing;
  const created = createParagraphWrapRegistry(flowDomainId);
  byFlowDomain.set(flowDomainId, created);
  return created;
}

export function commitOwnedParagraphWrapRegistry(
  owner: object,
  registry: ParagraphWrapRegistry,
  layout: ParagraphLayout,
): void {
  const byFlowDomain = registriesByOwner.get(owner);
  if (!byFlowDomain || byFlowDomain.get(registry.flowDomainId) !== registry) {
    throw new Error('Paragraph wrap registry transaction is stale');
  }
  byFlowDomain.set(
    registry.flowDomainId,
    commitParagraphWrapRegistry(registry, layout),
  );
}

/** Validate and retain one paragraph's occurrence-owned exclusion delta.
 * Returning a new value keeps acquisition transactional: failed candidates
 * cannot partially mutate the registry observed by later paragraphs. */
export function commitParagraphWrapRegistry(
  registry: ParagraphWrapRegistry,
  layout: ParagraphLayout,
): ParagraphWrapRegistry {
  if (layout.flowDomainId !== registry.flowDomainId) {
    throw new Error('Paragraph wrap registry cannot cross flow domains');
  }
  const known = new Set(registry.collisions.map((entry) => entry.occurrenceId));
  const collisionAdditions = ownedParagraphAnchorCollisions(layout);
  for (const entry of collisionAdditions) {
    if (known.has(entry.occurrenceId)) {
      throw new Error(`Paragraph wrap occurrence committed twice: ${entry.occurrenceId}`);
    }
    known.add(entry.occurrenceId);
  }
  const exclusionAdditions = ownedParagraphWrapExclusions(layout);
  const collisionAdditionIds = new Set(collisionAdditions.map((entry) =>
    entry.occurrenceId));
  const exclusionAdditionIds = new Set<string>();
  for (const exclusion of exclusionAdditions) {
    const occurrenceId = exclusion.anchorOccurrenceId;
    if (occurrenceId === undefined || !collisionAdditionIds.has(occurrenceId)) {
      throw new Error('Owned paragraph wrap exclusion omitted its collision occurrence');
    }
    if (exclusionAdditionIds.has(occurrenceId)) {
      throw new Error(`Paragraph wrap occurrence produced duplicate exclusions: ${occurrenceId}`);
    }
    exclusionAdditionIds.add(occurrenceId);
  }
  return Object.freeze({
    flowDomainId: registry.flowDomainId,
    collisions: Object.freeze([...registry.collisions, ...collisionAdditions]),
    exclusions: Object.freeze([...registry.exclusions, ...exclusionAdditions]),
  });
}
