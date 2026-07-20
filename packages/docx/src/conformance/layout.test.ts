/// <reference types="node" />

import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import type {
  DocParagraph,
  DocxDocumentModel,
  DocxTextRun,
  ImageRun,
} from '../types.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { createLayoutServices, layoutDocument } from '../renderer.js';
import { assertDocumentLayout, layoutFingerprint } from '../layout/invariants.js';
import type {
  DocumentLayout,
  LayoutRect,
  LayoutServices,
  ParagraphLayout,
  TextPlacement,
} from '../layout/types.js';
import {
  CONFORMANCE_CASES,
  coveredPairKeys,
  feasiblePairKeys,
} from './cases.js';
import { generateConformanceDocx, generateConformanceParts } from './generate.js';

function measureContext(): CanvasRenderingContext2D {
  return {
    font: '',
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText: (text: string) => ({
      width: [...text].length * 6,
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 8,
      fontBoundingBoxDescent: 2,
    }),
  } as unknown as CanvasRenderingContext2D;
}

function parse(bytes: Uint8Array): DocxDocumentModel {
  const archive = new DocxArchive(bytes);
  try {
    const parsed = JSON.parse(
      new TextDecoder().decode(archive.parse()),
    ) as DocxDocumentModel;
    return normalizeInternalDocumentModel(parsed).document;
  } finally {
    archive.free();
  }
}

function serviceIdentity(services: LayoutServices): readonly string[] {
  return [
    services.text.fingerprint,
    services.images.fingerprint,
    services.math.fingerprint,
  ];
}

function records(root: unknown): Record<string, unknown>[] {
  const output: Record<string, unknown>[] = [];
  const visit = (value: unknown): void => {
    if (value === null || typeof value !== 'object') return;
    if (Array.isArray(value)) {
      value.forEach(visit);
      return;
    }
    const record = value as Record<string, unknown>;
    output.push(record);
    Object.values(record).forEach(visit);
  };
  visit(root);
  return output;
}

function rootRecords(layout: DocumentLayout): Record<string, unknown>[] {
  return layout.pages.flatMap((page) =>
    page.layers.roots.flatMap(({ node }) => records(node)));
}

function paragraphRangePartitions(layout: DocumentLayout): void {
  const paragraphs = rootRecords(layout)
    .filter((record) => record.kind === 'paragraph') as unknown as ParagraphLayout[];
  expect(paragraphs.length).toBeGreaterThan(0);
  for (const paragraph of paragraphs) {
    let previousEnd = paragraph.lines[0]?.range.start ?? 0;
    for (const line of paragraph.lines) {
      expect(line.range.start).toBe(previousEnd);
      expect(line.range.end).toBeGreaterThanOrEqual(line.range.start);
      let coveredEnd = line.range.start;
      const sourceOrdered = [...line.placements]
        .sort((left, right) => left.range.start - right.range.start
          || left.range.end - right.range.end);
      for (const placement of sourceOrdered) {
        expect(placement.range.start).toBeGreaterThanOrEqual(line.range.start);
        expect(placement.range.end).toBeLessThanOrEqual(line.range.end);
        // RTL visual order and anchor-host markers may overlap source ranges,
        // but the union must remain a gap-free partition of the line.
        expect(placement.range.start).toBeLessThanOrEqual(coveredEnd);
        coveredEnd = Math.max(coveredEnd, placement.range.end);
      }
      expect(coveredEnd).toBe(line.range.end);
      previousEnd = line.range.end;
    }
  }
}

function targetTextPlacements(layout: DocumentLayout, target: string): TextPlacement[] {
  const paragraphs = rootRecords(layout)
    .filter((record) => record.kind === 'paragraph') as unknown as ParagraphLayout[];
  for (const paragraph of paragraphs) {
    const placements = paragraph.lines
      .flatMap(({ placements: linePlacements }) => linePlacements)
      .filter((placement): placement is TextPlacement => placement.kind === 'text')
      .sort((left, right) => left.range.start - right.range.start);
    if (placements.map(({ text }) => text).join('').includes(target)) return placements;
  }
  throw new Error(`retained text does not contain ${target}`);
}

function assertBottomClearance(layout: DocumentLayout): void {
  for (const page of layout.pages) {
    const domains = new Map(page.flowDomains.map((domain) => [domain.id, domain]));
    for (const { node } of page.layers.roots) {
      if (!node.ordinaryFlow) continue;
      const domain = domains.get(node.flowDomainId);
      expect(domain, `missing flow domain ${node.flowDomainId}`).toBeDefined();
      if (!domain) continue;
      const bounds = node.flowBounds as LayoutRect;
      const domainBottom = domain.logicalBounds.yPt + domain.logicalBounds.heightPt;
      const oversize = bounds.heightPt > domain.logicalBounds.heightPt;
      expect(
        bounds.yPt + bounds.heightPt <= domainBottom || oversize,
        `${node.id} invades the bottom of ${domain.id}`,
      ).toBe(true);
    }
  }
}

function maxTableDepth(value: unknown, depth = 0): number {
  if (value === null || typeof value !== 'object') return depth;
  if (Array.isArray(value)) {
    return value.reduce((max, entry) => Math.max(max, maxTableDepth(entry, depth)), depth);
  }
  const record = value as Record<string, unknown>;
  const nextDepth = record.type === 'table' ? depth + 1 : depth;
  return Object.values(record)
    .reduce<number>(
      (max, entry) => Math.max(max, maxTableDepth(entry, nextDepth)),
      nextDepth,
    );
}

function parsedDrawingCount(model: DocxDocumentModel): number {
  return records(model).filter((record) =>
    record.type === 'image' || record.type === 'shape').length;
}

function storyValue(
  model: DocxDocumentModel,
  story: 'body' | 'header' | 'footer',
): unknown {
  return story === 'body'
    ? model.body
    : story === 'header'
      ? model.headers.default
      : model.footers.default;
}

function targetParagraph(
  model: DocxDocumentModel,
  story: 'body' | 'header' | 'footer',
  target: string,
): DocParagraph {
  const paragraph = records(storyValue(model, story))
    .find((record) => record.type === 'paragraph'
      && JSON.stringify(record).includes(target));
  if (!paragraph) throw new Error(`parsed story does not contain ${target}`);
  return paragraph as unknown as DocParagraph;
}

function targetImage(
  model: DocxDocumentModel,
  story: 'body' | 'header' | 'footer',
): ImageRun | undefined {
  return records(storyValue(model, story))
    .find((record) => record.type === 'image') as unknown as ImageRun | undefined;
}

function textRunOf(paragraph: DocParagraph, target: string): DocxTextRun {
  const run = paragraph.runs.find((candidate) =>
    candidate.type === 'text' && candidate.text.includes(target));
  if (!run || run.type !== 'text') throw new Error(`parsed paragraph lacks ${target}`);
  return run;
}

function decodedPart(parts: ReadonlyMap<string, Uint8Array>, name: string): string {
  const bytes = parts.get(name);
  if (!bytes) throw new Error(`missing generated part ${name}`);
  return new TextDecoder().decode(bytes);
}

beforeAll(async () => {
  const wasm = await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url));
  await init({ module_or_path: wasm });
});

describe('synthetic DOCX conformance matrix', () => {
  it('is deterministic, bounded, and covers every feasible pair', () => {
    const feasible = [...feasiblePairKeys()].sort();
    const covered = [...coveredPairKeys(CONFORMANCE_CASES)].sort();
    expect(covered).toEqual(feasible);
    expect(CONFORMANCE_CASES.length).toBeGreaterThanOrEqual(12);
    expect(CONFORMANCE_CASES.length).toBeLessThanOrEqual(40);
    expect(new Set(CONFORMANCE_CASES.map(({ id }) => id)).size)
      .toBe(CONFORMANCE_CASES.length);
  });

  it('emits byte-identical stored ZIP archives on consecutive generations', () => {
    const first = CONFORMANCE_CASES.map(generateConformanceDocx);
    const second = CONFORMANCE_CASES.map(generateConformanceDocx);
    expect(second).toEqual(first);
    const hashes = (corpus: readonly Uint8Array[]) => corpus.map((bytes) =>
      createHash('sha256').update(bytes).digest('hex'));
    expect(hashes(second)).toEqual(hashes(first));
    for (const bytes of first) {
      expect([...bytes.subarray(0, 4)]).toEqual([0x50, 0x4b, 0x03, 0x04]);
    }
  });
});

describe.each(CONFORMANCE_CASES)('$id', (testCase) => {
  it('parses the intended OOXML facts through the real WASM parser', () => {
    const parts = generateConformanceParts(testCase);
    expect(parts.has('word/document.xml')).toBe(true);
    const model = parse(generateConformanceDocx(testCase));
    const paragraph = targetParagraph(
      model,
      testCase.expected.targetStory,
      testCase.expected.targetText,
    );
    const run = textRunOf(paragraph, testCase.expected.targetText);
    expect(Boolean(paragraph.bidi)).toBe(testCase.axes.direction === 'rtl');
    expect(paragraph.lineSpacing).toMatchObject({
      rule: testCase.axes.spacing,
      value: testCase.axes.spacing === 'exact' ? 12 : 1,
    });
    expect(Boolean(paragraph.lineSpacing?.explicit))
      .toBe(testCase.axes.styleSource !== 'documentDefault');
    expect(paragraph.styleId ?? null).toBe(
      testCase.axes.styleSource === 'paragraphStyle' ? 'CaseStyle' : 'Normal',
    );
    expect(run.fontFamily).toBe('Ahem');
    expect(maxTableDepth(model.body)).toBe(testCase.expected.tableDepth);
    expect(parsedDrawingCount(model)).toBe(testCase.expected.drawingCount);

    const image = targetImage(model, testCase.expected.targetStory);
    if (testCase.axes.object === 'none') {
      expect(image).toBeUndefined();
    } else {
      expect(image).toMatchObject({
        widthPt: 36,
        heightPt: 21.6,
        anchor: testCase.axes.object === 'floating',
      });
      if (image && testCase.axes.object === 'floating') {
        expect(Boolean(image.anchorXFromMargin))
          .toBe(testCase.axes.anchorReference !== 'page');
        expect(Boolean(image.anchorYFromPara))
          .toBe(testCase.axes.anchorReference === 'paragraph');
      }
    }

    // The parser proves the effective result above. These source-location
    // checks independently prove the generator placed the same value on the
    // requested OOXML inheritance layer, so the style/font axes cannot become
    // decorative while still producing an identical model.
    const storyPartName = testCase.axes.story === 'body'
      ? 'word/document.xml'
      : `word/${testCase.axes.story}1.xml`;
    const storyXml = decodedPart(parts, storyPartName);
    const stylesXml = decodedPart(parts, 'word/styles.xml');
    if (testCase.axes.styleSource === 'direct') {
      expect(storyXml).toContain('<w:spacing ');
    } else {
      expect(storyXml).not.toContain('<w:spacing ');
      expect(stylesXml).toContain('<w:spacing ');
    }
    if (testCase.axes.fontSource === 'direct') {
      expect(storyXml).toContain('w:ascii="Ahem"');
    } else if (testCase.axes.fontSource === 'theme') {
      expect(storyXml).toContain('w:asciiTheme="minorHAnsi"');
    } else {
      expect(storyXml).not.toContain('<w:rFonts');
      expect(stylesXml).toContain('w:ascii="Ahem"');
    }
  });

  it('retains deterministic, clone-safe geometry with complete source ranges', () => {
    const model = parse(generateConformanceDocx(testCase));
    const cloned = structuredClone(model);
    const firstServices = createLayoutServices(model, {
      measureContext: measureContext(),
    });
    const secondServices = createLayoutServices(cloned, {
      measureContext: measureContext(),
    });
    const first = layoutDocument(model, firstServices, { currentDateMs: 0 });
    const second = layoutDocument(cloned, secondServices, { currentDateMs: 0 });

    assertDocumentLayout(first);
    assertDocumentLayout(second);
    expect(structuredClone(first)).toEqual(first);
    expect(serviceIdentity(secondServices)).toEqual(serviceIdentity(firstServices));
    expect(layoutFingerprint(second)).toBe(layoutFingerprint(first));
    expect(first.pages).toHaveLength(testCase.expected.pageCount);
    expect(first.pages[0]?.geometry).toMatchObject({
      widthPt: testCase.expected.pageWidthPt,
      heightPt: testCase.expected.pageHeightPt,
    });
    expect(first.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);

    paragraphRangePartitions(first);
    const target = targetTextPlacements(first, testCase.expected.targetText);
    expect(target.length).toBeGreaterThan(0);
    for (const placement of target) {
      expect(placement.range.end).toBeGreaterThan(placement.range.start);
      expect([
        placement.bounds.xPt,
        placement.bounds.yPt,
        placement.bounds.widthPt,
        placement.bounds.heightPt,
      ].every(Number.isFinite)).toBe(true);
    }
    assertBottomClearance(first);
  });
});
