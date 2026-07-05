import { describe, it, expect } from 'vitest';
import {
  buildOutlineLayout,
  toggleGroupHidden,
  levelButtonHidden,
  type BandOutline,
} from './outline.js';

/** Build BandOutline[] from a compact `{index: [level, collapsed?, hidden?]}` map. */
function bands(spec: Record<number, [number, boolean?, boolean?]>): BandOutline[] {
  return Object.entries(spec).map(([i, [level, collapsed, hidden]]) => ({
    index: Number(i),
    level,
    collapsed: collapsed ?? false,
    hidden: hidden ?? false,
  }));
}

describe('buildOutlineLayout', () => {
  it('returns an empty layout when no band is grouped', () => {
    const layout = buildOutlineLayout(bands({ 1: [0], 2: [0] }), true);
    expect(layout.maxLevel).toBe(0);
    expect(layout.groups).toEqual([]);
  });

  it('places the summary AFTER the detail run when summaryBelow (default)', () => {
    // §18.3.1.73 example: rows 6-8 detail (levels 3,3,2), row 9 summary (level 1).
    const layout = buildOutlineLayout(
      bands({ 6: [3], 7: [3], 8: [2], 9: [1] }),
      true,
    );
    expect(layout.maxLevel).toBe(3);
    // Lane 1: the whole 6..9 run (all >= 1), summary is band 10.
    const l1 = layout.groups.find((g) => g.level === 1);
    expect(l1).toMatchObject({ level: 1, start: 6, end: 9, summary: 10 });
    // Lane 2: 6..8 (>= 2), summary is band 9.
    const l2 = layout.groups.find((g) => g.level === 2);
    expect(l2).toMatchObject({ level: 2, start: 6, end: 8, summary: 9 });
    // Lane 3: 6..7 (>= 3), summary is band 8.
    const l3 = layout.groups.find((g) => g.level === 3);
    expect(l3).toMatchObject({ level: 3, start: 6, end: 7, summary: 8 });
  });

  it('marks a group collapsed when its summary band carries collapsed', () => {
    // §18.3.1.73 "middle level collapsed": rows 6-8 hidden, row 9 (level 1)
    // has collapsed=1. The collapsed flag on band 9 describes its one-level-deeper
    // (level-2) children — i.e. the LANE-2 group (6..8) is collapsed, since band 9
    // is that group's summary. Lane 1 (summary band 10) is NOT collapsed.
    const layout = buildOutlineLayout(
      bands({ 6: [3, false, true], 7: [3, false, true], 8: [2, false, true], 9: [1, true] }),
      true,
    );
    const l2 = layout.groups.find((g) => g.level === 2);
    expect(l2?.summary).toBe(9);
    expect(l2?.collapsed).toBe(true);
    const l1 = layout.groups.find((g) => g.level === 1);
    expect(l1?.collapsed).toBe(false);
  });

  it('places the summary BEFORE the detail run when summaryAbove', () => {
    // Summary above: band 5 (level 1) is the summary for detail 6..7 (level 2).
    const layout = buildOutlineLayout(
      bands({ 5: [1], 6: [2], 7: [2] }),
      false,
    );
    const l2 = layout.groups.find((g) => g.level === 2);
    expect(l2).toMatchObject({ level: 2, start: 6, end: 7, summary: 5 });
  });

  it('yields a null summary when the run touches the sheet start (summaryAbove)', () => {
    const layout = buildOutlineLayout(bands({ 1: [1], 2: [1] }), false);
    const g = layout.groups.find((x) => x.level === 1);
    expect(g?.summary).toBeNull();
  });

  it('splits two separate groups at the same level', () => {
    const layout = buildOutlineLayout(
      bands({ 2: [1], 3: [1], 5: [1], 6: [1] }),
      true,
    );
    const l1 = layout.groups.filter((g) => g.level === 1);
    expect(l1).toHaveLength(2);
    expect(l1[0]).toMatchObject({ start: 2, end: 3, summary: 4 });
    expect(l1[1]).toMatchObject({ start: 5, end: 6, summary: 7 });
  });
});

describe('toggleGroupHidden', () => {
  const rowBands = bands({ 6: [3], 7: [3], 8: [2], 9: [1] });

  it('collapsing an expanded group hides its whole detail run', () => {
    const layout = buildOutlineLayout(rowBands, true);
    const l1 = layout.groups.find((g) => g.level === 1)!;
    const { hide, show, nowCollapsed } = toggleGroupHidden(l1, rowBands);
    expect(nowCollapsed).toBe(true);
    expect(hide).toEqual([6, 7, 8, 9]);
    expect(show).toEqual([]);
  });

  it('expanding a collapsed group reveals its detail run', () => {
    // "Middle level collapsed": lane-2 group (6..8) is collapsed (summary band 9
    // carries collapsed=1) and its detail rows 6-8 are hidden.
    const collapsedBands = bands({
      6: [3, false, true],
      7: [3, false, true],
      8: [2, false, true],
      9: [1, true, false],
    });
    const layout = buildOutlineLayout(collapsedBands, true);
    const l2 = layout.groups.find((g) => g.level === 2)!;
    expect(l2.collapsed).toBe(true);
    const { hide, show, nowCollapsed } = toggleGroupHidden(l2, collapsedBands);
    expect(nowCollapsed).toBe(false);
    expect(hide).toEqual([]);
    // Expanding reveals band 8 (the level-2 detail) directly under band 9.
    expect(show).toContain(8);
  });
});

describe('levelButtonHidden', () => {
  it('level 1 hides every grouped band', () => {
    const b = bands({ 6: [3], 7: [3], 8: [2], 9: [1] });
    const { hide, show } = levelButtonHidden(b, 1);
    expect(hide.sort()).toEqual([6, 7, 8, 9]);
    expect(show).toEqual([]);
  });

  it('level 2 hides bands at level >= 2, shows level-1 bands', () => {
    const b = bands({ 6: [3], 7: [3], 8: [2], 9: [1] });
    const { hide, show } = levelButtonHidden(b, 2);
    expect(hide.sort()).toEqual([6, 7, 8]);
    expect(show).toEqual([9]);
  });

  it('the maxLevel+1 button shows everything', () => {
    const b = bands({ 6: [3], 7: [3], 8: [2], 9: [1] });
    const { hide, show } = levelButtonHidden(b, 4);
    expect(hide).toEqual([]);
    expect(show.sort()).toEqual([6, 7, 8, 9]);
  });
});
