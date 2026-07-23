import type {
  DrawingMLCollisionEntryPt,
  DrawingMLCollisionRegistryDeltaPt,
  DrawingMLCollisionRegistrySnapshotPt,
  FloatRegistryCoordinateSpace,
} from './types.js';

function validateEntry(entry: DrawingMLCollisionEntryPt): void {
  if (entry.occurrenceId.length === 0) {
    throw new Error('DrawingML collision occurrence ID must not be empty');
  }
  const { xPt, yPt, widthPt, heightPt } = entry.bounds;
  if (![xPt, yPt, widthPt, heightPt].every(Number.isFinite)
    || widthPt < 0
    || heightPt < 0) {
    throw new Error(`DrawingML collision bounds are invalid: ${entry.occurrenceId}`);
  }
  if (
    (entry.horizontalOwnership !== 'page' && entry.horizontalOwnership !== 'host')
    || (entry.verticalOwnership !== 'page' && entry.verticalOwnership !== 'host')
  ) {
    throw new Error(`DrawingML collision ownership is invalid: ${entry.occurrenceId}`);
  }
}

function snapshotEntry(
  entry: DrawingMLCollisionEntryPt,
): DrawingMLCollisionEntryPt {
  validateEntry(entry);
  return Object.freeze({
    occurrenceId: entry.occurrenceId,
    bounds: Object.freeze({ ...entry.bounds }),
    horizontalOwnership: entry.horizontalOwnership,
    verticalOwnership: entry.verticalOwnership,
    ...(entry.relativeHeight !== undefined
      ? { relativeHeight: entry.relativeHeight }
      : {}),
  });
}

export function createDrawingMLCollisionRegistry(
  flowDomainId: string,
  coordinateSpace: FloatRegistryCoordinateSpace,
): DrawingMLCollisionRegistrySnapshotPt {
  return Object.freeze({
    coordinateSpace,
    flowDomainId,
    entries: Object.freeze([]),
  });
}

export function drawingMLCollisionRegistryDelta(
  snapshot: DrawingMLCollisionRegistrySnapshotPt,
  entries: readonly DrawingMLCollisionEntryPt[],
): DrawingMLCollisionRegistryDeltaPt {
  return Object.freeze({
    coordinateSpace: snapshot.coordinateSpace,
    flowDomainId: snapshot.flowDomainId,
    baseEntries: snapshot.entries,
    baseEntryCount: snapshot.entries.length,
    entries: Object.freeze(entries.map(snapshotEntry)),
  });
}

export function validateDrawingMLCollisionRegistryDelta(
  snapshot: DrawingMLCollisionRegistrySnapshotPt,
  delta: DrawingMLCollisionRegistryDeltaPt,
): void {
  if (delta.coordinateSpace !== snapshot.coordinateSpace) {
    throw new Error('DrawingML collision registry coordinate space mismatch');
  }
  if (delta.flowDomainId !== snapshot.flowDomainId) {
    throw new Error('DrawingML collision registry flow domain mismatch');
  }
  if (
    delta.baseEntries !== snapshot.entries
    || delta.baseEntryCount !== snapshot.entries.length
  ) {
    throw new Error('DrawingML collision registry delta is stale');
  }
  const known = new Set(snapshot.entries.map((entry) => entry.occurrenceId));
  for (const entry of delta.entries) {
    validateEntry(entry);
    if (known.has(entry.occurrenceId)) {
      throw new Error(`DrawingML collision occurrence committed twice: ${entry.occurrenceId}`);
    }
    known.add(entry.occurrenceId);
  }
}

export function applyDrawingMLCollisionRegistryDelta(
  snapshot: DrawingMLCollisionRegistrySnapshotPt,
  delta: DrawingMLCollisionRegistryDeltaPt,
): DrawingMLCollisionRegistrySnapshotPt {
  validateDrawingMLCollisionRegistryDelta(snapshot, delta);
  return Object.freeze({
    coordinateSpace: snapshot.coordinateSpace,
    flowDomainId: snapshot.flowDomainId,
    entries: Object.freeze([...snapshot.entries, ...delta.entries]),
  });
}
