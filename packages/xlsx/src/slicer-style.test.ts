import { describe, expect, it } from 'vitest';
import type { SlicerStyle } from './types';
import { drawSlicerFrame } from './renderer.js';

describe('custom slicer styles', () => {
  it('draws selected items with the resolved font, green rounded border, and no default frame', () => {
    const strokes: string[] = [];
    const texts: Array<{ text: string; fill: string; font: string }> = [];
    let roundedCorners = 0;
    const state: Record<string, unknown> = {
      fillStyle: '#000000',
      strokeStyle: '#000000',
      font: '10px sans-serif',
      textBaseline: 'alphabetic',
      textAlign: 'start',
      lineWidth: 1,
    };
    const ctx = new Proxy(state, {
      get(_target, prop: string) {
        if (prop in state) return state[prop];
        switch (prop) {
          case 'measureText':
            return (text: string) => ({ width: text.length * 7 });
          case 'quadraticCurveTo':
            return () => { roundedCorners += 1; };
          case 'stroke':
            return () => strokes.push(String(state.strokeStyle));
          case 'fillText':
            return (text: string) => texts.push({
              text,
              fill: String(state.fillStyle),
              font: String(state.font),
            });
          default:
            return () => undefined;
        }
      },
      set(_target, prop: string, value) {
        state[prop] = value;
        return true;
      },
    }) as unknown as CanvasRenderingContext2D;

    const style: SlicerStyle = {
      whole: {},
      header: {
        fontColor: '#5C7D21',
        fontSize: 12,
        fontBold: false,
        fontFamily: 'Meiryo UI',
      },
      selectedItemWithData: {
        fontColor: '#5C7D21',
        fontSize: 11,
        fontBold: false,
        fontFamily: 'Meiryo UI',
        borderColor: '#5C7D21',
      },
    };
    drawSlicerFrame(
      ctx,
      'Recipient',
      [{ name: 'Name 1', selected: true }],
      10,
      20,
      180,
      90,
      1,
      style,
    );

    expect(roundedCorners).toBe(4);
    expect(strokes).toEqual(['#5C7D21']);
    expect(texts).toEqual([
      { text: 'Recipient', fill: '#5C7D21', font: '12px "Meiryo UI", "Segoe UI", sans-serif' },
      { text: 'Name 1', fill: '#5C7D21', font: '11px "Meiryo UI", "Segoe UI", sans-serif' },
    ]);
  });
});
