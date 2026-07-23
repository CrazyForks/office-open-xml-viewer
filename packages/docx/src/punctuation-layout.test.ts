import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import { describe, expect, it } from 'vitest';
import {
  buildSegments,
  layoutLines,
  type LayoutLine,
  type LayoutTextSeg,
} from './line-layout.js';
import { createFontResolver } from './layout/font-service.js';
import {
  createTextLayoutService,
  type TextLayoutService,
} from './layout/text.js';
import type { LayoutServices } from './layout/types.js';
import type { DocParagraph, DocxTextRun } from './types.js';
import type { VerticalGlyphMeasurementService } from './layout/measurement-capabilities.js';

const VERTICAL_MEASUREMENT = {
  fingerprint: 'punctuation-layout:test',
  measureRunInkExtra: () => 0,
  planRun: () => [],
} as VerticalGlyphMeasurementService;

function context(): CanvasRenderingContext2D {
  let font = '10px serif';
  const size = () => Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    fontKerning: 'auto',
    letterSpacing: '0px',
    measureText: (text: string) => ({
      // Full-width synthetic face: exactly one em per scalar.
      width: [...text].length * size(),
      fontBoundingBoxAscent: size() * 0.8,
      fontBoundingBoxDescent: size() * 0.2,
      actualBoundingBoxAscent: size() * 0.8,
      actualBoundingBoxDescent: size() * 0.2,
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}

function textRun(text: string): DocParagraph['runs'][number] {
  const run: DocxTextRun = {
    text,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    fontSize: 10,
    color: null,
    fontFamily: 'Synthetic CJK',
    fontFamilyEastAsia: 'Synthetic CJK',
    isLink: false,
    background: null,
    vertAlign: null,
    hyperlink: null,
  };
  return {
    type: 'text',
    ...run,
    // Parser-only effective language input used by [MS-OE376] §2.1.56.
    langEastAsia: 'ja-jp',
  } as DocParagraph['runs'][number];
}

function lines(
  segments: ReturnType<typeof buildSegments>,
  width: number,
  overflowPunct: boolean,
): LayoutLine[] {
  return layoutLines(
    context(),
    segments,
    width,
    0,
    1,
    [],
    undefined,
    {},
    0,
    DEFAULT_KINSOKU_RULES,
    0,
    36,
    width,
    false,
    false,
    false,
    undefined,
    undefined,
    segments.some((segment) => 'text' in segment && segment.verticalRun)
      ? VERTICAL_MEASUREMENT
      : undefined,
    overflowPunct,
  );
}

const textOf = (line: LayoutLine): string =>
  line.segments
    .filter((segment): segment is LayoutTextSeg => 'text' in segment)
    .map((segment) => segment.text)
    .join('');

function punctuationMetricsServices(): LayoutServices {
  const text: TextLayoutService = createTextLayoutService({
    fonts: createFontResolver([]),
    measurer: {
      fingerprint: 'punctuation-layout:tight-ink',
      measure(request) {
        const scalarCount = [...request.text].length;
        return {
          advancePt: scalarCount * 10,
          ascentPt: 8,
          descentPt: 2,
          ...(request.text === '．'
            ? {
                inkBounds: { xMinPt: 1, xMaxPt: 3, ascentPt: 2, descentPt: 0 },
                horizontalInkBoundsAreTight: true,
              }
            : { inkBounds: { xMinPt: 0, xMaxPt: scalarCount * 10, ascentPt: 8, descentPt: 2 } }),
        };
      },
    },
  });
  return { text } as unknown as LayoutServices;
}

describe('ECMA-376 East-Asian punctuation fit', () => {
  it('compresses a full-width trailing period by half an em in measure and paint geometry', () => {
    const segments = buildSegments([textRun('甲乙．')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
    });

    expect(segments.map((segment) => 'text' in segment ? segment.text : '')).toEqual([
      '甲乙',
      '．',
    ]);
    const punctuation = segments[1] as LayoutTextSeg;
    expect(punctuation.charSpacing).toBeUndefined();
    expect(punctuation.punctuationCompressionPt).toBe(-5);

    const laidOut = lines(segments, 25, false);
    expect(laidOut).toHaveLength(1);
    expect(textOf(laidOut[0])).toBe('甲乙．');
    expect((laidOut[0].segments[1] as LayoutTextSeg).measuredWidth).toBe(5);
  });

  it('trims the selected punctuation glyph to its tight trailing ink edge', () => {
    const segments = buildSegments([textRun('甲乙．')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
      layoutServices: punctuationMetricsServices(),
    });

    const punctuation = segments[1] as LayoutTextSeg;
    expect(punctuation.text).toBe('．');
    expect(punctuation.punctuationCompressionPt).toBe(-7);
    const laidOut = lines(segments, 23, false);
    expect(laidOut).toHaveLength(1);
    expect(textOf(laidOut[0])).toBe('甲乙．');
    expect((laidOut[0].segments[1] as LayoutTextSeg).measuredWidth).toBe(3);
  });

  it('scales punctuation sidebearing trim with the authored glyph width', () => {
    const run = {
      ...textRun('．'),
      charScale: 0.5,
    } as DocParagraph['runs'][number];
    const [punctuation] = buildSegments([run], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
      layoutServices: punctuationMetricsServices(),
    }) as LayoutTextSeg[];

    // Natural 10pt cell and 7pt trailing sidebearing are both scaled to 50%.
    expect(punctuation.punctuationCompressionPt).toBe(-3.5);
    expect(lines([punctuation], 1.5, false)).toHaveLength(1);
  });

  it('allows one eligible punctuation character past the line extent by default policy', () => {
    const segments = buildSegments([textRun('甲乙．')], {
      pageIndex: 0,
      totalPages: 1,
    });

    const hanging = lines(segments, 20, true);
    expect(hanging).toHaveLength(1);
    expect(textOf(hanging[0])).toBe('甲乙．');

    const disabled = lines(buildSegments([textRun('甲乙．')], {
      pageIndex: 0,
      totalPages: 1,
    }), 20, false);
    expect(disabled).toHaveLength(2);
    expect(disabled.map(textOf)).toEqual(['甲', '乙．']);
  });

  it('does not uniformly compress a Japanese middle dot', () => {
    const segments = buildSegments([textRun('甲乙・丙丁')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
    });

    expect(segments).toHaveLength(1);
    expect((segments[0] as LayoutTextSeg).text).toBe('甲乙・丙丁');
    expect((segments[0] as LayoutTextSeg).punctuationCompressionPt).toBeUndefined();
    const [line] = lines(segments, 100, false);
    expect((line.segments[0] as LayoutTextSeg).measuredWidth).toBe(50);
  });

  it('preserves cross-segment kinsoku ownership when compression splits the punctuation', () => {
    const segments = buildSegments([textRun('一二三四五六、七')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
    });

    expect(segments.map((segment) => 'text' in segment ? segment.text : '')).toEqual([
      '一二三四五六',
      '、',
      '七',
    ]);
    const wrapped = lines(segments, 60, false);
    expect(wrapped.map(textOf)).toEqual(['一二三四五', '六、七']);
  });

  it('applies the document-wide punctuation compression in vertical CJK columns', () => {
    const text = `${'甲'.repeat(37)}。`;
    const compressed = buildSegments([textRun(text)], {
      pageIndex: 0,
      totalPages: 1,
      verticalCJK: true,
      characterSpacingControl: 'compressPunctuation',
    });
    // 37 full-em cells + one compressed half-em punctuation cell = 375pt.
    expect(lines(compressed, 375, false)).toHaveLength(1);
    const compressedMark = compressed.at(-1) as LayoutTextSeg;
    expect(compressedMark.text).toBe('。');
    expect(compressedMark.verticalRun).toBe(true);
    expect(compressedMark.punctuationCompressionPt).toBe(-5);

    const uncompressed = buildSegments([textRun(text)], {
      pageIndex: 0,
      totalPages: 1,
      verticalCJK: true,
      characterSpacingControl: 'doNotCompress',
    });
    expect(lines(uncompressed, 375, false)).toHaveLength(2);

    const middleDot = buildSegments([textRun(`${'甲'.repeat(37)}・`)], {
      pageIndex: 0,
      totalPages: 1,
      verticalCJK: true,
      characterSpacingControl: 'compressPunctuation',
    });
    expect(middleDot).toHaveLength(1);
    expect((middleDot[0] as LayoutTextSeg).punctuationCompressionPt).toBeUndefined();
    expect(lines(middleDot, 375, false)).toHaveLength(2);
  });
});
