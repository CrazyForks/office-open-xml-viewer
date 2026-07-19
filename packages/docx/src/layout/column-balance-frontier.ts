import { adjustForWidowOrphan } from '../line-fit-policy.js';
import type { BodyLayoutInput } from './body-layout-input.js';
import {
  solveExactColumnBalance,
  type ColumnBalanceFragment,
} from './column-balancing.js';
import type { BodyFlowAllocation } from './section-flow-composition.js';
import { sourceKey } from './source-key.js';
import type { LayoutPage, PageSectionRegion } from './types.js';

interface RetainedBalanceFragment extends ColumnBalanceFragment {
  readonly sequenceIndex: number;
  readonly sourceIdentity: string;
  readonly paragraphLine: boolean;
}

function balanceSourcePolicies(input: BodyLayoutInput): ReadonlyMap<string, Readonly<{
  sequenceIndex: number;
  keepLines: boolean;
  keepNext: boolean;
  widowControl: boolean;
}>> {
  const policies = new Map<string, Readonly<{
    sequenceIndex: number;
    keepLines: boolean;
    keepNext: boolean;
    widowControl: boolean;
  }>>();
  input.sequence.forEach((entry, sequenceIndex) => {
    if (entry.kind === 'body-block') {
      const block = entry.block;
      policies.set(sourceKey(block.source), Object.freeze({
        sequenceIndex,
        keepLines: block.kind === 'paragraph' && block.keepLines,
        keepNext: block.kind === 'paragraph' && block.keepNext,
        widowControl: block.kind === 'paragraph' && block.widowControl,
      }));
      return;
    }
    if (entry.kind !== 'adjacent-table-group') return;
    entry.tables.forEach((table) => policies.set(sourceKey(table.source), Object.freeze({
      sequenceIndex,
      keepLines: false,
      keepNext: false,
      widowControl: false,
    })));
  });
  return policies;
}

function splitRetainedExtent(
  blockStartPt: number,
  blockEndPt: number,
  retainedBoundariesPt: readonly number[],
): readonly number[] {
  const extentPt = blockEndPt - blockStartPt;
  if (retainedBoundariesPt.length <= 1) return Object.freeze([extentPt]);
  const result: number[] = [];
  let priorBoundaryPt = blockStartPt;
  for (const boundaryPt of retainedBoundariesPt) {
    if (!Number.isFinite(boundaryPt)
      || boundaryPt < priorBoundaryPt
      || boundaryPt > blockEndPt) return Object.freeze([extentPt]);
    result.push(boundaryPt - priorBoundaryPt);
    priorBoundaryPt = boundaryPt;
  }
  // The final fragment owns trailing paragraph spacing/table flow charge. This
  // derives from the accepted allocation rather than recreating spacing policy.
  result[result.length - 1] = result.at(-1)! + blockEndPt - priorBoundaryPt;
  return Object.freeze(result);
}

function retainedColumnBalanceFragments(
  input: BodyLayoutInput,
  allocations: readonly BodyFlowAllocation[],
  page: LayoutPage,
  outgoing: PageSectionRegion,
): readonly RetainedBalanceFragment[] {
  const domains = new Set(outgoing.flowDomainIds);
  const nodes = new Map(page.layers.body.map((node) => [node.id, node]));
  const policies = balanceSourcePolicies(input);
  const fragments: Array<{
    extentPt: number;
    breakAfter: ColumnBalanceFragment['breakAfter'];
    sequenceIndex: number;
    sourceIdentity: string;
    paragraphLine: boolean;
  }> = [];
  for (const allocation of allocations) {
    if (!domains.has(allocation.flowDomainId)) continue;
    const node = nodes.get(allocation.nodeId);
    if (!node || !node.ordinaryFlow) continue;
    const identity = sourceKey(node.source);
    const policy = policies.get(identity);
    if (!policy) continue;
    const retainedBoundariesPt = node.kind === 'paragraph'
      && !policy.keepLines
      && node.lines.length > 1
      ? node.lines.map((line) => line.bounds.yPt + line.advancePt)
      : node.kind === 'table' && node.rows.length > 1
        ? node.rows.map((row) => row.flowBounds.yPt + row.advancePt)
        : [allocation.blockEndPt];
    const pieces = splitRetainedExtent(
      allocation.blockStartPt,
      allocation.blockEndPt,
      retainedBoundariesPt,
    );
    pieces.forEach((piece) => fragments.push({
      extentPt: piece,
      breakAfter: 'allowed',
      sequenceIndex: policy.sequenceIndex,
      sourceIdentity: identity,
      paragraphLine: node.kind === 'paragraph'
        && !policy.keepLines
        && node.lines.length > 0,
    }));
  }

  const bySource = new Map<string, number[]>();
  fragments.forEach((fragment, index) => {
    const indexes = bySource.get(fragment.sourceIdentity) ?? [];
    indexes.push(index);
    bySource.set(fragment.sourceIdentity, indexes);
  });
  for (const [identity, indexes] of bySource) {
    const policy = policies.get(identity);
    if (!policy) continue;
    const lineIndexes = indexes.filter((index) => fragments[index]!.paragraphLine);
    if (policy.keepLines) {
      indexes.slice(0, -1).forEach((index) => {
        fragments[index]!.breakAfter = 'forbidden';
      });
    }
    for (let ordinal = 0; ordinal + 1 < lineIndexes.length; ordinal += 1) {
      if (adjustForWidowOrphan({
        widowControl: policy.widowControl,
        start: 0,
        end: ordinal + 1,
        totalLines: lineIndexes.length,
        canRelocate: true,
      }).kind === 'keep') continue;
      fragments[lineIndexes[ordinal]!]!.breakAfter = 'forbidden';
    }
    if (policy.keepNext) {
      fragments[indexes.at(-1)!]!.breakAfter = 'forbidden';
    }
  }
  let activeSectionOccurrenceId = input.initialSection.sectionOccurrenceId;
  input.sequence.forEach((entry, sequenceIndex) => {
    if (entry.kind === 'begin-section') {
      activeSectionOccurrenceId = entry.section.sectionOccurrenceId;
      return;
    }
    if (activeSectionOccurrenceId !== outgoing.sectionOccurrenceId
      || entry.kind !== 'authored-break'
      || entry.break !== 'column') return;
    for (let priorIndex = fragments.length - 1; priorIndex >= 0; priorIndex -= 1) {
      if (fragments[priorIndex]!.sequenceIndex >= sequenceIndex) continue;
      fragments[priorIndex]!.breakAfter = 'forced';
      break;
    }
  });
  return Object.freeze(fragments.map((fragment) => Object.freeze(fragment)));
}

export function exactRetainedColumnBalanceTarget(
  input: BodyLayoutInput,
  allocations: readonly BodyFlowAllocation[],
  footnoteReserveByPage: ReadonlyMap<number, number>,
  page: LayoutPage,
  outgoing: PageSectionRegion,
): number {
  const fragments = retainedColumnBalanceFragments(input, allocations, page, outgoing);
  const reservePt = footnoteReserveByPage.get(page.pageIndex) ?? 0;
  return solveExactColumnBalance({
    columnCount: outgoing.flowDomainIds.length,
    fragments,
  }).targetPt + reservePt;
}
