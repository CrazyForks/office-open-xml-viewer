import { describe, expect, it } from 'vitest';
import viewer from '../viewer.ts?raw';

describe('XLSX viewer composition boundary', () => {
  it('the composite facade subclasses the engine while the canvas facade owns exactly one engine', () => {
    expect(viewer).toContain("super(container, opts, { kind: 'composite' })");
    expect(viewer.match(/new XlsxViewerEngine\(/g)).toHaveLength(1);
    expect(viewer).not.toContain("parseSheetLocally");
  });
});
