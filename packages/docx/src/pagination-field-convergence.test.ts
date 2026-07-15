import { describe, expect, it } from 'vitest';
import {
  convergePaginationFields,
  paginationFieldFlowGeometry,
  paginationFieldGeometryFingerprint,
} from './layout/pagination-fields.js';

describe('pagination field convergence seam', () => {
  it('normalizes absent optional runtime placement facts before fingerprinting', () => {
    const omitted = paginationFieldGeometryFingerprint({
      pageCount: 1,
      pages: [[{ type: 'paragraph', colIndex: 0 }]],
    });
    const explicitUndefined = paginationFieldGeometryFingerprint({
      pageCount: 1,
      pages: [[{ type: 'paragraph', colIndex: 0, colY: undefined, placed: undefined }]],
    });

    expect(explicitUndefined).toBe(omitted);
  });

  it('projects paragraph geometry without parser/source objects', () => {
    const geometry = paginationFieldFlowGeometry({
      kind: 'paragraph', id: 'body:0',
      source: { story: 'body', storyInstance: 'body', path: [0] },
      flowDomainId: 'body', ordinaryFlow: true,
      flowBounds: { xPt: 10, yPt: 20, widthPt: 100, heightPt: 12 },
      inkBounds: { xPt: 10, yPt: 20, widthPt: 40, heightPt: 10 },
      advancePt: 12, spacing: { beforePt: 0, afterPt: 2 }, contextualSpacing: false,
      lines: [], borders: [], resources: [], drawings: [], textBoxes: [], events: [], exclusions: [],
    });

    expect(geometry).toMatchObject({
      kind: 'paragraph',
      flowBounds: { xPt: 10, yPt: 20, widthPt: 100, heightPt: 12 },
      advancePt: 12,
    });
    expect(JSON.stringify(geometry)).not.toContain('source');
    expect(JSON.stringify(geometry)).not.toContain('storyInstance');
  });

  it('stabilizes on the geometry acquired with the resolved page count', () => {
    const hints: number[] = [];
    const result = convergePaginationFields((hint) => {
      hints.push(hint);
      const pageCount = hint === 1 ? 2 : 2;
      return { fingerprint: `pages:${pageCount}`, pageCount };
    });

    expect(hints).toEqual([1, 2]);
    expect(result).toEqual({ fingerprint: 'pages:2', pageCount: 2 });
  });

  it('hard-fails a repeated geometry cycle', () => {
    let step = 0;
    expect(() => convergePaginationFields(() => {
      const current = step++ % 2 === 0
        ? { fingerprint: 'geometry:a', pageCount: 2 }
        : { fingerprint: 'geometry:b', pageCount: 1 };
      return current;
    })).toThrow(/repeated geometry fingerprint cycle/i);
  });

  it('hard-fails when hostile geometry never stabilizes within the policy limit', () => {
    let step = 0;
    expect(() => convergePaginationFields(
      () => ({ fingerprint: `geometry:${step++}`, pageCount: step + 1 }),
      3,
    )).toThrow(/hard iteration limit 3 reached/i);
  });
});
