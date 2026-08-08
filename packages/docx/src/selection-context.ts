import type { TextSelectionContextOptions } from '@silurus/ooxml-core';
import { readBoundedNativeTextSelection } from '@silurus/ooxml-core/internal/canvas-viewer-mechanics';

/** Bounds for a DOCX selection-context snapshot. Extensible per format. */
export type DocxSelectionContextOptions = TextSelectionContextOptions;

/** Snapshot-local locator for a rendered run intersecting a DOCX text selection. */
export interface DocxSelectionRunLocator {
  readonly pageIndex: number;
  readonly runIndex: number;
  readonly paragraphId?: string;
}

/** Detached, bounded context for a read-only AI/MCP handoff. */
export interface DocxSelectionContext {
  readonly format: 'docx';
  readonly kind: 'text';
  readonly text: string;
  readonly pageIndexes: readonly number[];
  readonly paragraphIds: readonly string[];
  readonly runs: readonly DocxSelectionRunLocator[];
  readonly truncated: boolean;
  readonly truncationReasons: readonly ('text' | 'runs')[];
  readonly textCharacters: number;
  readonly maxTextCharacters: number;
  readonly maxRunLocators: number;
}

function nonNegativeInteger(value: string | undefined): number | null {
  if (value === undefined || !/^\d+$/.test(value)) return null;
  const number = Number(value);
  return Number.isSafeInteger(number) ? number : null;
}

function pageIndexFor(run: HTMLElement): number | null {
  for (let element: HTMLElement | null = run; element; element = element.parentElement) {
    const index = nonNegativeInteger(element.dataset.pageIndex);
    if (index !== null) return index;
  }
  return null;
}

export function readDocxSelectionContext(
  root: HTMLElement,
  selection: Selection | null,
  options: DocxSelectionContextOptions = {},
): DocxSelectionContext | null {
  const bounded = readBoundedNativeTextSelection(root, selection, (run) => {
    const pageIndex = pageIndexFor(run);
    const runIndex = nonNegativeInteger(run.dataset.runIndex);
    if (pageIndex === null || runIndex === null) return null;
    return {
      pageIndex,
      runIndex,
      ...(run.dataset.paragraphId === undefined ? {} : { paragraphId: run.dataset.paragraphId }),
    } satisfies DocxSelectionRunLocator;
  }, {
    maxChars: options.maxTextCharacters,
    maxLocators: options.maxRunLocators,
  });
  if (!bounded) return null;
  const runs = [...bounded.locators].sort(
    (left, right) => left.pageIndex - right.pageIndex || left.runIndex - right.runIndex,
  );
  return {
    format: 'docx',
    kind: 'text',
    text: bounded.text,
    pageIndexes: [...new Set(runs.map((run) => run.pageIndex))],
    paragraphIds: [...new Set(runs.flatMap((run) => run.paragraphId ? [run.paragraphId] : []))],
    runs,
    truncated: bounded.truncated,
    truncationReasons: bounded.truncationReasons,
    textCharacters: bounded.textCharacters,
    maxTextCharacters: bounded.maxTextCharacters,
    maxRunLocators: bounded.maxLocators,
  };
}
