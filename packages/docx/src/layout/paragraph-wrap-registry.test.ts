import { describe, expect, it } from 'vitest';
import {
  accessParagraphWrapRegistry,
  commitOwnedParagraphWrapRegistry,
  commitParagraphWrapRegistry,
  createParagraphWrapRegistry,
  inheritedParagraphAuthorityForReacquisition,
  ownedParagraphAnchorCollisions,
  ownedParagraphWrapExclusions,
} from './paragraph-wrap-registry.js';
import type {
  DrawingMLCollisionEntryPt,
  ParagraphLayout,
  WrapExclusion,
} from './types.js';

function exclusion(occurrenceId: string): WrapExclusion {
  return {
    id: `exclusion:${occurrenceId}`,
    wrap: 'square',
    bounds: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
    polygon: [
      { xPt: 10, yPt: 20 },
      { xPt: 40, yPt: 20 },
      { xPt: 40, yPt: 60 },
      { xPt: 10, yPt: 60 },
    ],
    anchorOccurrenceId: occurrenceId,
    verticalOwnership: 'host',
  };
}

function paragraph(
  flowDomainId: string,
  ownedOccurrenceIds: readonly string[],
  exclusions: readonly WrapExclusion[],
): ParagraphLayout {
  const collisions: DrawingMLCollisionEntryPt[] = [
    {
      occurrenceId: 'prior',
      bounds: { xPt: 0, yPt: 0, widthPt: 5, heightPt: 5 },
      horizontalOwnership: 'host',
      verticalOwnership: 'host',
    },
    ...ownedOccurrenceIds.map((occurrenceId) => ({
      occurrenceId,
      bounds: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
      horizontalOwnership: 'host' as const,
      verticalOwnership: 'host' as const,
    })),
  ];
  return {
    flowDomainId,
    exclusions,
    anchorCollisions: collisions,
    anchorFrames: ownedOccurrenceIds.map((occurrenceId) => ({
      status: 'resolved',
      occurrenceId,
    })),
    drawings: ownedOccurrenceIds.map((occurrenceId) => ({
      anchorLayer: {
        occurrenceId,
        behindDoc: false,
        relativeHeight: 0,
        sourceOrder: 0,
        horizontalOwnership: 'host',
        verticalOwnership: 'host',
      },
    })),
  } as unknown as ParagraphLayout;
}

describe('paragraph wrap occurrence registry', () => {
  it('keeps owner-local flow registries transactional without renderer state', () => {
    const owner = {};
    const registry = accessParagraphWrapRegistry(owner, 'cell:0');
    const layout = paragraph('cell:0', ['current'], [exclusion('current')]);

    commitOwnedParagraphWrapRegistry(owner, registry, layout);

    expect(accessParagraphWrapRegistry(owner, 'cell:0').exclusions)
      .toEqual([exclusion('current')]);
    expect(() => commitOwnedParagraphWrapRegistry(owner, registry, layout))
      .toThrow(/stale/);
    expect(accessParagraphWrapRegistry({}, 'cell:0').exclusions).toEqual([]);
  });

  it('returns a new immutable snapshot containing only exclusions owned by the paragraph', () => {
    const inherited = exclusion('prior');
    const owned = exclusion('current');
    const initial = createParagraphWrapRegistry('cell:0');
    const committed = commitParagraphWrapRegistry(
      initial,
      paragraph('cell:0', ['current'], [inherited, owned]),
    );

    expect(ownedParagraphWrapExclusions(
      paragraph('cell:0', ['current'], [inherited, owned]),
    )).toEqual([owned]);
    expect(ownedParagraphAnchorCollisions(
      paragraph('cell:0', ['current'], [inherited, owned]),
    ).map((entry) => entry.occurrenceId)).toEqual(['current']);
    expect(committed).not.toBe(initial);
    expect(initial.exclusions).toEqual([]);
    expect(initial.collisions).toEqual([]);
    expect(committed.exclusions).toEqual([owned]);
    expect(committed.collisions.map((entry) => entry.occurrenceId)).toEqual(['current']);
    expect(Object.isFrozen(committed)).toBe(true);
    expect(Object.isFrozen(committed.exclusions)).toBe(true);
    expect(Object.isFrozen(committed.collisions)).toBe(true);
  });

  it('fails closed when a paragraph crosses flow domains', () => {
    const registry = createParagraphWrapRegistry('cell:0');

    expect(() => commitParagraphWrapRegistry(
      registry,
      paragraph('cell:1', ['current'], [exclusion('current')]),
    )).toThrow(/flow domains/);
  });

  it('fails closed when an owned occurrence was already committed', () => {
    const owned = exclusion('current');
    const first = commitParagraphWrapRegistry(
      createParagraphWrapRegistry('cell:0'),
      paragraph('cell:0', ['current'], [owned]),
    );

    expect(() => commitParagraphWrapRegistry(
      first,
      paragraph('cell:0', ['current'], [owned]),
    )).toThrow(/committed twice/);
  });

  it('indexes a wrapNone occurrence by collision geometry without an exclusion', () => {
    const committed = commitParagraphWrapRegistry(
      createParagraphWrapRegistry('cell:0'),
      paragraph('cell:0', ['current'], []),
    );

    expect(committed.collisions.map((entry) => entry.occurrenceId)).toEqual(['current']);
    expect(committed.exclusions).toEqual([]);
  });

  it('fails closed when one owned occurrence produces duplicate exclusions', () => {
    const owned = exclusion('current');

    expect(() => commitParagraphWrapRegistry(
      createParagraphWrapRegistry('cell:0'),
      paragraph('cell:0', ['current'], [owned, { ...owned, id: 'duplicate' }]),
    )).toThrow(/duplicate exclusions/);
  });

  it('preserves persistent inherited authority while replacing final-frame probes', () => {
    const current = paragraph('cell:0', ['current'], [
      {
        ...exclusion('persistent-table'),
        id: 'persistent-table',
        anchorOccurrenceId: undefined,
      },
      {
        ...exclusion('transient-table'),
        id: 'table-final-frame:0',
        anchorOccurrenceId: undefined,
      },
      exclusion('prior'),
      exclusion('current'),
    ]);
    const inherited = inheritedParagraphAuthorityForReacquisition(current);

    expect(inherited.exclusions.map((entry) => entry.id)).toEqual([
      'persistent-table',
      'exclusion:prior',
    ]);
    expect(inherited.collisions.map((entry) => entry.occurrenceId)).toEqual(['prior']);
    expect(inheritedParagraphAuthorityForReacquisition({
      ...current,
      exclusions: inherited.exclusions,
      anchorCollisions: inherited.collisions,
    } as ParagraphLayout)).toEqual(inherited);
  });
});
