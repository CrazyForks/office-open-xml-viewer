// Regression tests for two chart data-fidelity fixes:
//
//   1. Scatter connecting line: a series-level `<c:spPr><a:ln><a:noFill/>`
//      (ECMA-376 §21.2.2.198) OVERRIDES the chart-group `<c:scatterStyle>`
//      (§21.2.2.42). Excel draws a MARKERS-ONLY scatter when the series line is
//      `<a:noFill/>` even though the group default is `lineMarker`. The parser
//      records this as `ChartSeries.lineHidden`; the renderer must draw no
//      connecting line for such a series (sample-30 sheet 1's two scatters).
//
//   2. Pie per-point data labels: a per-point `<c:dLbl idx>` (§21.2.2.47)
//      carries its own show-flag group (§21.2.2.177/.180/.187/.189) and text
//      style, which OVERRIDE the series-level `<c:dLbls>` defaults (§21.2.2.49)
//      for that one slice. sample-14 slide-7's pie sets `showCatName=0
//      showPercent=1` + white text PER SLICE while the series default is
//      `showCatName=1` black — so each label must render as WHITE PERCENT-ONLY,
//      not black "category + percent".
//
// Both broke when the pptx/xlsx parsers were unified onto the shared
// ooxml-common `parse_chart_part` (which now surfaces scatterStyle /
// series-level dLbls the renderer honors) without carrying the override /
// noFill semantics that suppress or reshape those defaults.

import { describe, it, expect } from 'vitest';
import type { ChartModel, ChartSeries, ChartRect } from '../types/chart';
import { renderChart } from './renderer.js';

const RECT: ChartRect = { x: 0, y: 0, w: 640, h: 360 };

function baseModel(over: Partial<ChartModel>): ChartModel {
  return {
    chartType: 'clusteredBar',
    title: null,
    categories: [],
    series: [],
    showDataLabels: false,
    valMin: null,
    valMax: null,
    catAxisTitle: null,
    valAxisTitle: null,
    catAxisHidden: false,
    valAxisHidden: false,
    catAxisLineHidden: false,
    valAxisLineHidden: false,
    plotAreaBg: null,
    chartBg: null,
    showLegend: false,
    legendPos: null,
    catAxisCrossBetween: 'between',
    valAxisMajorTickMark: 'out',
    catAxisMajorTickMark: 'out',
    titleFontSizeHpt: null,
    titleFontColor: null,
    titleFontFace: null,
    catAxisFontSizeHpt: null,
    valAxisFontSizeHpt: null,
    dataLabelFontSizeHpt: null,
    subtotalIndices: [],
    ...over,
  };
}

function series(over: Partial<ChartSeries>): ChartSeries {
  return { name: '', color: null, values: [], ...over };
}

interface TextCall { text: string; fillStyle: string }

/** Recording context that captures: (a) `stroke()` calls that follow a
 *  `moveTo`/`lineTo` path with ≥2 vertices AND whose strokeStyle is `matchColor`
 *  (the series line color) — this isolates the scatter connecting line from
 *  gridlines/axis rules, which stroke in grey; (b) `fillText` calls with the
 *  fillStyle in effect. Enough to assert the scatter line pass and pie labels. */
function recordingCtx(matchColor = '#4f81bd'): {
  ctx: CanvasRenderingContext2D;
  counts: { polylineStrokes: number };
  texts: TextCall[];
} {
  const texts: TextCall[] = [];
  const counts = { polylineStrokes: 0 };
  let pathVerts = 0;
  const state: Record<string, unknown> = {
    font: '10px sans-serif',
    fillStyle: '#000',
    strokeStyle: '#000',
    lineWidth: 1,
    textAlign: 'start',
    textBaseline: 'alphabetic',
    lineCap: 'butt',
    lineJoin: 'miter',
    globalAlpha: 1,
  };
  const fontPx = (font: string): number => {
    const m = /(\d+(?:\.\d+)?)px/.exec(font);
    return m ? parseFloat(m[1]) : 10;
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText':
          return (t: string) => {
            const px = fontPx(String(state.font));
            let w = 0;
            for (const ch of String(t)) w += ch.charCodeAt(0) > 0x2e7f ? px : px * 0.6;
            return { width: w };
          };
        case 'beginPath':
          return () => { pathVerts = 0; };
        case 'moveTo':
        case 'lineTo':
        case 'bezierCurveTo':
        case 'quadraticCurveTo':
          return () => { pathVerts += 1; };
        // A marker is a single `arc` + `fill`; it does not build a multi-vertex
        // moveTo/lineTo path. Only a stroke in the SERIES color with ≥2 vertices
        // is a connecting line (gridlines/axes stroke grey → excluded).
        case 'stroke':
          return () => {
            if (pathVerts >= 2 && String(state.strokeStyle).toLowerCase() === matchColor.toLowerCase()) {
              counts.polylineStrokes += 1;
            }
          };
        case 'fillText':
          return (text: string) =>
            texts.push({ text, fillStyle: String(state.fillStyle) });
        case 'createLinearGradient':
        case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        default:
          return () => undefined;
      }
    },
    set(_t, prop: string, value) { state[prop] = value; return true; },
  };
  return {
    ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D,
    counts,
    texts,
  };
}

describe('scatter series-line noFill overrides scatterStyle (§21.2.2.198)', () => {
  // sample-30 sheet 1: `<c:scatterStyle val="lineMarker">` at the group level,
  // but each series carries `<c:spPr><a:ln><a:noFill/></c:spPr>` → markers only.
  const scatterModel = (lineHidden: boolean | null): ChartModel =>
    baseModel({
      chartType: 'scatter',
      scatterStyle: 'lineMarker',
      categories: ['1', '2', '3', '4'],
      series: [series({
        name: 'S',
        color: '4f81bd',
        values: [10, 20, 15, 25],
        categories: ['1', '2', '3', '4'],
        showMarker: true,
        lineHidden,
      })],
    });

  it('draws NO connecting polyline when the series line is noFill', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, scatterModel(true), RECT, 1);
    expect(rec.counts.polylineStrokes).toBe(0);
  });

  it('still draws the connecting polyline for a normal lineMarker scatter', () => {
    // Guard: the fix must not suppress lines universally — a series WITHOUT the
    // noFill flag keeps the group scatterStyle line (regression the other way).
    const rec = recordingCtx();
    renderChart(rec.ctx, scatterModel(null), RECT, 1);
    expect(rec.counts.polylineStrokes).toBeGreaterThan(0);
  });
});

describe('pie per-point dLbl flags override series defaults (§21.2.2.47)', () => {
  // sample-14 slide-7 pie: series default `<c:dLbls>` says showCatName=1,
  // black; every slice's `<c:dLbl>` says showCatName=0 showPercent=1, WHITE.
  // Result must be white percent-only labels.
  const pieModel = (): ChartModel =>
    baseModel({
      chartType: 'pie',
      showDataLabels: true,
      categories: ['Alpha', 'Beta', 'Gamma'],
      series: [series({
        name: 'Share',
        values: [50, 30, 20],
        categories: ['Alpha', 'Beta', 'Gamma'],
        dataPointColors: ['ff0000', '00ff00', '0000ff'],
        seriesDataLabels: {
          showVal: false,
          showCatName: true,   // series default WOULD show category names…
          showSerName: false,
          showPercent: true,
          fontColor: '000000', // …in black.
          fontBold: true,
          fontSizeHpt: 1800,
          formatCode: '0%',
          position: 'ctr',
        },
        dataLabelOverrides: [0, 1, 2].map(idx => ({
          idx,
          text: '',            // no custom text → compose from flags
          fontColor: 'FFFFFF', // per-slice WHITE
          fontBold: true,
          fontSizeHpt: 1200,
          showVal: false,
          showCatName: false,  // per-slice: percent only
          showSerName: false,
          showPercent: true,
        })),
      })],
    });

  it('renders white, percent-only labels (no category names, not black)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, pieModel(), RECT, 1);
    const labels = rec.texts.filter(t => /%$/.test(t.text) || /Alpha|Beta|Gamma/.test(t.text));
    expect(labels.length).toBeGreaterThan(0);
    // Every drawn slice label is a bare percentage…
    for (const l of labels) {
      expect(l.text).toMatch(/^\d+%$/);
      // …in white (the per-point override color), never the series-default black.
      expect(l.fillStyle.toLowerCase()).toBe('#ffffff');
    }
    // And no category name leaked through.
    expect(rec.texts.some(t => /Alpha|Beta|Gamma/.test(t.text))).toBe(false);
  });

  it('a genuine <c:delete> per-point override still skips that slice', () => {
    const model = pieModel();
    model.series[0].dataLabelOverrides = [
      { idx: 0, text: '', deleted: true },
      ...([1, 2].map(idx => ({
        idx, text: '', fontColor: 'FFFFFF', showCatName: false, showPercent: true,
      }))),
    ];
    const rec = recordingCtx();
    renderChart(rec.ctx, model, RECT, 1);
    const pctLabels = rec.texts.filter(t => /^\d+%$/.test(t.text));
    // Two slices labeled (idx 1,2); the deleted idx 0 is skipped.
    expect(pctLabels.length).toBe(2);
  });
});
