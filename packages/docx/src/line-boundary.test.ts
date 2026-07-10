import { describe, expect, it } from 'vitest';
import {
  DEFAULT_KINSOKU_RULES,
  type KinsokuRules,
} from '@silurus/ooxml-core';
import {
  layoutLines,
  type LayoutLine,
  type LayoutSeg,
  type LayoutTextSeg,
  type LineBoundary,
} from './line-layout.js';
import type { TabStop } from './types.js';

function makeLinearCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const fontSize = (): number => Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (text: string) => {
      const size = fontSize();
      return {
        width: [...text].length * size * 0.5,
        fontBoundingBoxAscent: size * 0.8,
        fontBoundingBoxDescent: size * 0.2,
        actualBoundingBoxAscent: size * 0.8,
        actualBoundingBoxDescent: size * 0.2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

function textSeg(text: string, extra: Partial<LayoutTextSeg> = {}): LayoutTextSeg {
  return {
    text,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    fontSize: 10,
    color: null,
    fontFamily: 'Times New Roman',
    vertAlign: null,
    measuredWidth: 0,
    ...extra,
  };
}

const lineBreak = (): LayoutSeg => ({ lineBreak: true, fontSize: 10, measuredWidth: 0 });

interface Fixture {
  readonly name: string;
  readonly width: number;
  readonly segs: () => LayoutSeg[];
  readonly tabStops?: TabStop[];
  readonly kinsoku?: KinsokuRules;
  readonly hasTrailingBreak?: boolean;
}

const crossRunKinsoku: KinsokuRules = {
  enabled: true,
  lineStartForbidden: new Set(['。'.codePointAt(0)!]),
  lineEndForbidden: new Set(),
};

const fixtures: Fixture[] = [
  {
    name: 'plain Latin words',
    width: 35,
    segs: () => ['alpha ', 'bravo ', 'charlie ', 'delta ', 'echo ', 'foxtrot '].map((text) => textSeg(text)),
  },
  {
    name: 'CJK per-glyph split',
    width: 25,
    segs: () => [textSeg('あ'.repeat(18))],
  },
  {
    name: 'kinsoku cross-run retraction',
    width: 20,
    segs: () => [textSeg('あいうえ'), textSeg('。かきくけこさしす')],
    kinsoku: crossRunKinsoku,
  },
  {
    name: 'manual breaks including a trailing break',
    width: 200,
    segs: () => [textSeg('first'), lineBreak(), textSeg('second'), lineBreak()],
    hasTrailingBreak: true,
  },
  {
    name: 'right-aligned custom tab with committed trailing text',
    width: 60,
    tabStops: [{ pos: 40, alignment: 'right', leader: 'dot' }],
    segs: () => [
      textSeg('name'),
      { isTab: true, fontSize: 10, measuredWidth: 0 },
      textSeg('99'),
      lineBreak(),
      textSeg('tail '),
      textSeg('words '),
      textSeg('wrap'),
    ],
  },
  {
    name: 'over-long Latin overflow-wrap',
    width: 20,
    segs: () => [textSeg('abcdefghijklmnopqrst')],
  },
  {
    name: 'small-caps glued group',
    width: 55,
    segs: () => [
      textSeg('Lead '),
      textSeg('I', { smallCaps: true }),
      textSeg('NTRODUCTION', { smallCaps: true, joinPrev: true }),
      textSeg(' tail '),
      textSeg('words'),
    ],
  },
];

function layoutFixture(fixture: Fixture, startBoundary?: LineBoundary): LayoutLine[] {
  const segs = fixture.segs().map((seg) => ({ ...seg })) as LayoutSeg[];
  return layoutLines(
    makeLinearCtx(),
    segs,
    fixture.width,
    0,
    1,
    fixture.tabStops,
    undefined,
    {},
    0,
    fixture.kinsoku ?? DEFAULT_KINSOKU_RULES,
    0,
    36,
    fixture.width,
    false,
    startBoundary,
  );
}

function textSequence(lines: LayoutLine[]): string[] {
  return lines.map((line) => line.segments
    .filter((segment): segment is LayoutTextSeg => 'text' in segment)
    .map((segment) => segment.text)
    .join(''));
}

function isEnd(boundary: LineBoundary, segCount: number): boolean {
  return boundary.segIndex === segCount && boundary.charOffset === 0;
}

describe('layoutLines consumed-content boundaries', () => {
  for (const fixture of fixtures) {
    it(`reproduces every content-bearing suffix for ${fixture.name}`, () => {
      const lines = layoutFixture(fixture);
      expect(lines.length).toBeGreaterThan(0);

      for (let i = 0; i < lines.length; i++) {
        const boundary = lines[i].consumedEnd;
        expect(boundary).toBeDefined();
        if (!boundary) continue;

        const expected = textSequence(lines.slice(i + 1));
        const suffix = textSequence(layoutFixture(fixture, boundary));

        // A trailing manual break creates a final zero-content line. Its
        // predecessor and that empty line both necessarily end at the same
        // two-field END coordinate, so END cannot distinguish those two entry
        // states. The mid-stream boundary below still proves that suffix layout
        // preserves and re-emits the trailing empty line.
        if (fixture.hasTrailingBreak && isEnd(boundary, fixture.segs().length) && expected.length > 0) {
          expect(expected).toEqual(['']);
          expect(suffix).toEqual([]);
          continue;
        }
        expect(suffix).toEqual(expected);
      }
    });
  }

  it('places a cross-run kinsoku boundary inside the retracted source segment', () => {
    const fixture = fixtures.find((candidate) => candidate.name === 'kinsoku cross-run retraction')!;
    const lines = layoutFixture(fixture);

    expect(lines.some((line) =>
      line.consumedEnd?.segIndex === 0 && line.consumedEnd.charOffset > 0,
    )).toBe(true);
  });

  it('re-emits a trailing-break empty line from an earlier original-stream boundary', () => {
    const fixture = fixtures.find((candidate) => candidate.hasTrailingBreak)!;
    const lines = layoutFixture(fixture);
    const boundary = lines[0].consumedEnd;

    expect(boundary).toBeDefined();
    expect(textSequence(lines)).toEqual(['first', 'second', '']);
    expect(textSequence(layoutFixture(fixture, boundary))).toEqual(['second', '']);
  });

  for (const fixture of fixtures) {
    it(`records monotonic boundaries ending at END for ${fixture.name}`, () => {
      const segCount = fixture.segs().length;
      const boundaries = layoutFixture(fixture).map((line) => line.consumedEnd);
      expect(boundaries.every((boundary) => boundary !== undefined)).toBe(true);

      for (let i = 1; i < boundaries.length; i++) {
        const previous = boundaries[i - 1]!;
        const current = boundaries[i]!;
        expect(
          current.segIndex > previous.segIndex
            || (current.segIndex === previous.segIndex && current.charOffset >= previous.charOffset),
        ).toBe(true);
      }
      expect(boundaries.at(-1)).toEqual({ segIndex: segCount, charOffset: 0 });
    });
  }
});
