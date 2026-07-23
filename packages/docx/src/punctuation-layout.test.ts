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
import { wordIsOverflowPunctuation } from './layout/line-compatibility.js';
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

function textRun(
  text: string,
  eastAsiaLanguage = 'ja-jp',
): DocParagraph['runs'][number] {
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
    // Parser-only effective language input consumed by the isolated
    // overflow-punctuation compatibility projection.
    langEastAsia: eastAsiaLanguage,
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
        const tightTrailingWhitespace = new Set([
          '。', '．', '！', '？', '：', '；', 'あ', 'ア', 'ー', 'か\u3099',
        ]).has(request.text);
        const advancePt = scalarCount * 10;
        return {
          advancePt,
          ascentPt: 8,
          descentPt: 2,
          ...(tightTrailingWhitespace
            ? {
                inkBounds: {
                  xMinPt: 1,
                  xMaxPt: request.text === '．' ? 3 : advancePt - 5,
                  ascentPt: 2,
                  descentPt: 0,
                },
                horizontalInkBoundsAreTight: true,
              }
            : { inkBounds: { xMinPt: 0, xMaxPt: advancePt, ascentPt: 8, descentPt: 2 } }),
        };
      },
    },
  });
  return { text } as unknown as LayoutServices;
}

describe('ECMA-376 East-Asian punctuation fit', () => {
  it('compresses a full-width trailing period from tight selected-glyph geometry', () => {
    const segments = buildSegments([textRun('甲乙．')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
      layoutServices: punctuationMetricsServices(),
    });

    expect(segments.map((segment) => 'text' in segment ? segment.text : '')).toEqual([
      '甲乙',
      '．',
    ]);
    const punctuation = segments[1] as LayoutTextSeg;
    expect(punctuation.charSpacing).toBeUndefined();
    expect(punctuation.punctuationCompressionPt).toBe(-7);

    const laidOut = lines(segments, 23, false);
    expect(laidOut).toHaveLength(1);
    expect(textOf(laidOut[0])).toBe('甲乙．');
    expect((laidOut[0].segments[1] as LayoutTextSeg).measuredWidth).toBe(3);
  });

  it('does not invent a fixed trim when tight horizontal ink bounds are unavailable', () => {
    const segments = buildSegments([textRun('甲乙．')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
    });

    const punctuation = segments[1] as LayoutTextSeg;
    expect(punctuation.text).toBe('．');
    expect(punctuation.punctuationCompressionPt).toBeUndefined();
    expect(lines(segments, 23, false)).toHaveLength(2);
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

  it('dispatches exact ST_CharacterSpacing values and compresses kana only in the combined mode', () => {
    const punctuationOnly = buildSegments([textRun('あアー．')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
      layoutServices: punctuationMetricsServices(),
    }) as LayoutTextSeg[];
    expect(punctuationOnly.map((segment) => segment.text)).toEqual(['あアー', '．']);
    expect(punctuationOnly.map((segment) => segment.punctuationCompressionPt)).toEqual([
      undefined,
      -7,
    ]);

    const punctuationAndKana = buildSegments([textRun('あアー・漢．')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuationAndJapaneseKana',
      layoutServices: punctuationMetricsServices(),
    }) as LayoutTextSeg[];
    expect(punctuationAndKana.map((segment) => segment.text)).toEqual([
      'あ',
      'ア',
      'ー',
      '・漢',
      '．',
    ]);
    expect(punctuationAndKana.map((segment) => segment.punctuationCompressionPt)).toEqual([
      -5,
      -5,
      -5,
      undefined,
      -7,
    ]);

    const unknownPrefix = buildSegments([textRun('あ．')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuationFuture',
      layoutServices: punctuationMetricsServices(),
    }) as LayoutTextSeg[];
    expect(unknownPrefix).toHaveLength(1);
    expect(unknownPrefix[0].punctuationCompressionPt).toBeUndefined();
  });

  it('does not compress halfwidth punctuation as full-width punctuation', () => {
    const segments = buildSegments([textRun('｡､')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
    }) as LayoutTextSeg[];

    expect(segments).toHaveLength(1);
    expect(segments[0].text).toBe('｡､');
    expect(segments[0].punctuationCompressionPt).toBeUndefined();
  });

  it('keeps a decomposed kana base and combining mark in one compressed grapheme', () => {
    const decomposedGa = 'か\u3099';
    const segments = buildSegments([textRun(`甲${decomposedGa}乙`)], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuationAndJapaneseKana',
      layoutServices: punctuationMetricsServices(),
    }) as LayoutTextSeg[];

    expect(segments.map((segment) => segment.text)).toEqual(['甲', decomposedGa, '乙']);
    expect(segments[1].punctuationCompressionPt).toBe(-5);
    expect(segments.every((segment) => segment.text !== '\u3099')).toBe(true);
  });

  it('recognizes unambiguous full-width exclamation, question, colon, and semicolon punctuation', () => {
    const segments = buildSegments([textRun('！甲？乙：丙；')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
      layoutServices: punctuationMetricsServices(),
    }) as LayoutTextSeg[];

    expect(segments.map((segment) => segment.text)).toEqual([
      '！', '甲', '？', '乙', '：', '丙', '；',
    ]);
    expect(segments.map((segment) => segment.punctuationCompressionPt)).toEqual([
      -5, undefined, -5, undefined, -5, undefined, -5,
    ]);
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
      layoutServices: punctuationMetricsServices(),
    });
    // 37 full-em cells + the selected period's retained 5pt ink extent = 375pt.
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

  it('admits an eligible trailing punctuation character independently of script', () => {
    const latin = buildSegments([textRun('A B.', 'en-us')], {
      pageIndex: 0,
      totalPages: 1,
    });

    expect(lines(latin, 30, true).map(textOf)).toEqual(['A B.']);
    expect(lines(buildSegments([textRun('A B.', 'en-us')], {
      pageIndex: 0,
      totalPages: 1,
    }), 30, false).map(textOf)).toEqual(['A ', 'B.']);
  });

  it('finds the final visible punctuation before a collapsible separator', () => {
    const enabled = buildSegments([textRun('A B. C', 'en-us')], {
      pageIndex: 0,
      totalPages: 1,
    });
    const disabled = buildSegments([textRun('A B. C', 'en-us')], {
      pageIndex: 0,
      totalPages: 1,
    });

    expect(lines(enabled, 30, true).map(textOf)).toEqual(['A B. ', 'C']);
    expect(lines(disabled, 30, false).map(textOf)).toEqual(['A ', 'B. ', 'C']);
  });

  it('limits U+3017 to the Simplified Chinese overflow-punctuation set', () => {
    expect(wordIsOverflowPunctuation('〗', 'zh-cn')).toBe(true);
    expect(wordIsOverflowPunctuation('〗', 'zh-hans')).toBe(true);
    expect(wordIsOverflowPunctuation('〗', 'zh-tw')).toBe(false);
    expect(wordIsOverflowPunctuation('〗', 'zh-hant')).toBe(false);
    expect(wordIsOverflowPunctuation('〗', 'zh-mo')).toBe(false);
    expect(wordIsOverflowPunctuation('〗', 'ja-jp')).toBe(false);
    expect(wordIsOverflowPunctuation('〗', 'ko-kr')).toBe(false);
  });
});
