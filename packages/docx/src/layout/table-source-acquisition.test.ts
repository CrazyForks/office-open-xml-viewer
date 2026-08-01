import { describe, expect, it } from 'vitest';
import {
  tableColumnLayoutInput,
  tableSourceAcquisitionInput,
} from '../parser-model.js';
import type { DocTable } from '../types.js';
import { projectTableColumnLayoutInput } from './table-source-acquisition.js';

const noBorders = Object.freeze({
  top: null, right: null, bottom: null, left: null, insideH: null, insideV: null,
});

function sourceTable(): DocTable {
  return {
    colWidths: [36],
    rows: [{
      gridBefore: 1,
      gridAfter: 1,
      rowHeight: null,
      rowHeightRule: 'auto',
      isHeader: false,
      cells: [{
        content: [],
        colSpan: 2,
        vMerge: null,
        borders: noBorders,
        background: null,
        vAlign: 'top',
        widthPt: null,
        widthPct: 2500,
        __tableCellLayout: {
          preferredWidth: { kind: 'pct', value: '3000' },
          margins: null,
        },
      }],
      __tableRowLayout: {
        height: null,
        justification: null,
        beforeWidth: { kind: 'pct', value: '10%' },
        afterWidth: { kind: 'dxa', value: '200' },
        cellSpacing: null,
        exception: null,
      },
    }],
    borders: noBorders,
    cellMarginTop: 0,
    cellMarginRight: 0,
    cellMarginBottom: 0,
    cellMarginLeft: 0,
    jc: 'left',
    layout: 'autofit',
    widthPct: 3750,
    __tableLayout: {
      effectiveStyleId: 'Synthetic',
      ordinaryFlow: true,
      grid: {
        authored: true,
        columns: [{ width: '720' }],
        requiredColumnCount: 4,
      },
      preferredWidth: { kind: 'pct', value: '3750' },
      layout: { kind: 'autofit' },
      cellSpacing: null,
    },
  } as unknown as DocTable;
}

describe('table source acquisition input', () => {
  it('keeps historical hand-built rich-textbox tables without colWidths compatible', () => {
    const table = sourceTable();
    delete (table as unknown as { colWidths?: number[] }).colWidths;
    delete (table as unknown as { __tableLayout?: unknown }).__tableLayout;

    const input = tableSourceAcquisitionInput(table as DocTable);

    expect(input.semantic.colWidths).toEqual([]);
    expect(tableColumnLayoutInput(
      table as DocTable,
      200,
      () => ({ minWidthPt: 12, maxWidthPt: 30 }),
    ).gridWidthsPt).toEqual([]);
  });

  it('retains only frozen clone-safe column semantics and no cell contents or parser wire', () => {
    const table = sourceTable();
    table.rows[0]!.cells[0]!.content.push({
      type: 'paragraph',
      runs: [{ type: 'text', text: 'must-not-be-retained' }],
    } as never);

    const input = tableSourceAcquisitionInput(table);

    expect(input.semantic.rows[0]!.cells[0]).toEqual({
      colSpan: 2,
      widthPt: null,
      widthPct: 2500,
    });
    expect(JSON.stringify(input.semantic)).not.toContain('must-not-be-retained');
    expect(JSON.stringify(input.semantic)).not.toContain('__table');
    expect(Object.isFrozen(input)).toBe(true);
    expect(Object.isFrozen(input.semantic)).toBe(true);
    expect(Object.isFrozen(input.semantic.rows[0]!.cells[0])).toBe(true);
    expect(() => structuredClone(input)).not.toThrow();
  });

  it('does not duplicate a deeply nested table subtree in the outer fact', () => {
    const outer = sourceTable();
    let current = outer.rows[0]!.cells[0]!;
    for (let depth = 0; depth < 20; depth += 1) {
      const nested = sourceTable();
      current.content.push({ type: 'table', ...nested });
      current = nested.rows[0]!.cells[0]!;
    }
    current.content.push({
      type: 'paragraph',
      runs: [{ type: 'text', text: 'deep-content-sentinel'.repeat(1_000) }],
    } as never);

    const serialized = JSON.stringify(tableSourceAcquisitionInput(outer));

    expect(serialized).not.toContain('deep-content-sentinel');
    expect(serialized.length).toBeLessThan(2_000);
  });

  it('delegates the compatibility API to the deterministic pure projector with exact parity', () => {
    const table = sourceTable();
    const cell = table.rows[0]!.cells[0]!;
    const compatibility = tableColumnLayoutInput(table, 200, (received) => {
      expect(received).toBe(cell);
      return { minWidthPt: 12, maxWidthPt: 30 };
    }, 240);
    const input = tableSourceAcquisitionInput(table);
    const project = () => projectTableColumnLayoutInput(
      input,
      200,
      (rowIndex, cellIndex) => {
        expect([rowIndex, cellIndex]).toEqual([0, 0]);
        return { minWidthPt: 12, maxWidthPt: 30 };
      },
      240,
    );

    expect(project()).toEqual(compatibility);
    expect(project()).toEqual(project());
    expect(tableSourceAcquisitionInput(table)).toBe(input);
  });
});
