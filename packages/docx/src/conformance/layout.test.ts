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
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocument } from '../document-layout.js';
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
import {
  generateConformanceDocx,
  generateConformanceParts,
  storeZip,
} from './generate.js';

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
    expect(paragraph.lines.length, `${paragraph.id} has no retained lines`)
      .toBeGreaterThan(0);
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

function encodedXml(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

function inlineDrawingCase() {
  const testCase = CONFORMANCE_CASES.find(({ axes }) =>
    axes.story === 'body'
    && axes.container === 'paragraph'
    && axes.object === 'inline');
  if (!testCase) throw new Error('conformance corpus lacks a body inline-drawing case');
  return testCase;
}

function duplicateInlineDrawingParts(
  kind: 'image' | 'chart',
): Map<string, Uint8Array> {
  const parts = new Map(generateConformanceParts(inlineDrawingCase()));
  const documentXml = decodedPart(parts, 'word/document.xml');
  const drawing = documentXml.match(
    /<w:r><w:drawing>[\s\S]*?<\/w:drawing><\/w:r>/u,
  )?.[0];
  if (!drawing) throw new Error('generated inline drawing is unavailable');
  const relationships = decodedPart(parts, 'word/_rels/document.xml.rels');

  if (kind === 'image') {
    const missing = drawing.replaceAll('rIdImage', 'rIdMissingImage');
    const available = drawing
      .replaceAll('rIdImage', 'rIdAvailableImage')
      .replaceAll('id="1"', 'id="2"');
    parts.set(
      'word/document.xml',
      encodedXml(documentXml.replace(drawing, `${missing}${available}`)),
    );
    parts.set(
      'word/_rels/document.xml.rels',
      encodedXml(relationships.replace(
        /<Relationship Id="rIdImage"[^>]*\/>/u,
        '<Relationship Id="rIdMissingImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/missing.png"/>'
        + '<Relationship Id="rIdAvailableImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/pixel.png"/>',
      )),
    );
    return parts;
  }

  const graphic = (relationshipId: string) =>
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">`
    + '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
    + `<c:chart r:id="${relationshipId}"/>`
    + '</a:graphicData></a:graphic>';
  const chartDrawing = (relationshipId: string, id: string) => drawing
    .replace(/<a:graphic[\s\S]*<\/a:graphic>/u, graphic(relationshipId))
    .replaceAll('id="1"', `id="${id}"`);
  parts.set(
    'word/document.xml',
    encodedXml(documentXml.replace(
      drawing,
      chartDrawing('rIdMissingChart', '3')
      + chartDrawing('rIdAvailableChart', '4'),
    )),
  );
  parts.set(
    'word/_rels/document.xml.rels',
    encodedXml(relationships.replace(
      /<Relationship Id="rIdImage"[^>]*\/>/u,
      '<Relationship Id="rIdMissingChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/missing.xml"/>'
      + '<Relationship Id="rIdAvailableChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>',
    )),
  );
  parts.set(
    '[Content_Types].xml',
    encodedXml(decodedPart(parts, '[Content_Types].xml').replace(
      '</Types>',
      '<Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>',
    )),
  );
  parts.delete('word/media/pixel.png');
  parts.set('word/charts/chart1.xml', encodedXml(
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">'
    + '<c:chart><c:plotArea><c:barChart><c:barDir val="col"/>'
    + '<c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/>'
    + '<c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt>'
    + '</c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>',
  ));
  return parts;
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
    expect(feasible).toContain('story=header|fontSource=theme');
    expect(feasible).not.toContain('story=header|container=table');
    expect(feasible).toContain('object=floating|anchorReference=paragraph');
    expect(feasible).not.toContain('object=inline|anchorReference=page');
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

describe('recoverable missing drawing resources', () => {
  it.each(['image', 'chart'] as const)(
    'keeps a valid %s addressable after a missing drawing in authored order',
    (kind) => {
      const model = parse(storeZip(duplicateInlineDrawingParts(kind)));
      const runs = (model.body[0] as DocParagraph).runs;
      expect(runs.map((run) => run.type)).toEqual([
        'text',
        'image',
        kind,
      ]);
      expect(runs[1]).toMatchObject({
        type: 'image',
        unavailableResourceKind: kind,
        imagePath: '',
      });

      const services = createLayoutServices(model, { measureContext: measureContext() });
      const layout = layoutDocument(model, services, { currentDateMs: 0 });

      assertDocumentLayout(layout);
      expect(layout.diagnostics).toEqual(expect.arrayContaining([
        expect.objectContaining({ code: 'MISSING_RESOURCE', severity: 'warning' }),
      ]));
      expect(rootRecords(layout)).toEqual(expect.arrayContaining([
        expect.objectContaining({
          kind: 'resource',
          resourceKind: kind,
        }),
      ]));
    },
  );

  it('retains inline geometry and visible content through the real WASM-to-layout pipeline', () => {
    const testCase = inlineDrawingCase();
    const parts = new Map(generateConformanceParts(testCase));
    parts.delete('word/media/pixel.png');

    const model = parse(storeZip(parts));
    expect(records(model).some((record) =>
      record.type === 'image'
      && record.unavailableResourceKind === 'image')).toBe(true);
    expect(JSON.stringify(model)).not.toContain('unavailableDrawing');
    expect(JSON.stringify(structuredClone(model))).not.toContain('__anchorAcquisition');

    const services = createLayoutServices(model, { measureContext: measureContext() });
    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const cloned = structuredClone(model);
    const clonedServices = createLayoutServices(cloned, { measureContext: measureContext() });
    const clonedLayout = layoutDocument(cloned, clonedServices, { currentDateMs: 0 });

    assertDocumentLayout(layout);
    assertDocumentLayout(clonedLayout);
    expect(serviceIdentity(clonedServices)).toEqual(serviceIdentity(services));
    expect(layoutFingerprint(clonedLayout)).toBe(layoutFingerprint(layout));
    expect(targetTextPlacements(layout, testCase.expected.targetText).length)
      .toBeGreaterThan(0);
    expect(layout.diagnostics).toEqual(expect.arrayContaining([
      expect.objectContaining({
        code: 'MISSING_RESOURCE',
        severity: 'warning',
      }),
    ]));
    expect(rootRecords(layout)).toEqual(expect.arrayContaining([
      expect.objectContaining({
        kind: 'drawing',
        commands: [{ kind: 'noop' }],
        flowBounds: expect.objectContaining({ widthPt: 36, heightPt: 21.6 }),
      }),
    ]));
  });

  it('retains anchored recovery geometry after the public model is cloned', () => {
    const testCase = CONFORMANCE_CASES.find(({ axes }) =>
      axes.story === 'body'
      && axes.container === 'paragraph'
      && axes.object === 'floating');
    if (!testCase) throw new Error('conformance corpus lacks a body floating-drawing case');
    const parts = new Map(generateConformanceParts(testCase));
    parts.delete('word/media/pixel.png');

    const model = parse(storeZip(parts));
    const cloned = structuredClone(model);
    expect(records(cloned).some((record) =>
      record.type === 'image'
      && record.unavailableResourceKind === 'image')).toBe(true);
    expect(JSON.stringify(cloned)).not.toContain('unavailableDrawing');
    expect(JSON.stringify(cloned)).not.toContain('__anchorAcquisition');

    const services = createLayoutServices(model, { measureContext: measureContext() });
    const clonedServices = createLayoutServices(cloned, { measureContext: measureContext() });
    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const clonedLayout = layoutDocument(cloned, clonedServices, { currentDateMs: 0 });

    assertDocumentLayout(layout);
    assertDocumentLayout(clonedLayout);
    expect(serviceIdentity(clonedServices)).toEqual(serviceIdentity(services));
    expect(layoutFingerprint(clonedLayout)).toBe(layoutFingerprint(layout));
    expect(layout.diagnostics).toEqual(expect.arrayContaining([
      expect.objectContaining({ code: 'MISSING_RESOURCE', severity: 'warning' }),
    ]));
    expect(rootRecords(layout)).toEqual(expect.arrayContaining([
      expect.objectContaining({
        kind: 'drawing',
        commands: [{ kind: 'noop' }],
      }),
    ]));
  });
});
