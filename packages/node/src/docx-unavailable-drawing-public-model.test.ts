import { describe, expect, it } from 'vitest';
import { CONFORMANCE_CASES } from '../../docx/src/conformance/cases.ts';
import {
  generateConformanceParts,
  storeZip,
} from '../../docx/src/conformance/generate.ts';
import { parseDocx } from './docx.ts';

function objectRecords(value: unknown): Array<Record<string, unknown>> {
  const output: Array<Record<string, unknown>> = [];
  const visit = (candidate: unknown): void => {
    if (candidate === null || typeof candidate !== 'object') return;
    if (Array.isArray(candidate)) {
      candidate.forEach(visit);
      return;
    }
    const record = candidate as Record<string, unknown>;
    output.push(record);
    Object.values(record).forEach(visit);
  };
  visit(value);
  return output;
}

describe('Node DOCX public parser projection', () => {
  it('returns a declared clone-safe recovery run without parser-private fields', () => {
    const testCase = CONFORMANCE_CASES.find(({ axes }) =>
      axes.story === 'body'
      && axes.container === 'paragraph'
      && axes.object === 'inline');
    if (!testCase) throw new Error('conformance corpus lacks an inline drawing case');
    const parts = new Map(generateConformanceParts(testCase));
    parts.delete('word/media/pixel.png');

    const model = parseDocx(storeZip(parts));
    const cloned = structuredClone(model);
    const recovery = objectRecords(cloned).find((record) =>
      record.type === 'image' && record.unavailableResourceKind === 'image');

    expect(recovery).toMatchObject({
      type: 'image',
      imagePath: '',
      mimeType: '',
      unavailableResourceKind: 'image',
      widthPt: 36,
      heightPt: 21.6,
    });
    expect(JSON.stringify(cloned)).not.toContain('__anchorAcquisition');
    expect(JSON.stringify(cloned)).not.toContain('unavailableDrawing');
  });
});
