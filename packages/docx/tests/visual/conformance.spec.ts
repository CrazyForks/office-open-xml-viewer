import { expect, test } from '@playwright/test';
import { CONFORMANCE_CASES } from '../../src/conformance/cases.js';

interface BrowserConformanceReport {
  readonly caseId: string;
  readonly expected: (typeof CONFORMANCE_CASES)[number]['expected'];
  readonly pageCount: number;
  readonly mainSizes: readonly { widthPt: number; heightPt: number }[];
  readonly workerSizes: readonly { widthPt: number; heightPt: number }[];
  readonly runsFinite: boolean;
  readonly targetTextPresent: boolean;
  readonly cloneSafe: boolean;
  readonly layoutFingerprint: string;
  readonly mainWorkerParity: {
    readonly pageCount: boolean;
    readonly sizes: boolean;
    readonly runs: boolean;
  };
  readonly authoredGeometry: {
    readonly pageBoxes: readonly {
      readonly widthPt: number;
      readonly heightPt: number;
      readonly contentTopPt: number;
      readonly contentBottomPt: number;
    }[];
    readonly tables: readonly {
      readonly id: string;
      readonly columnWidthsPt: readonly number[];
      readonly borderEndpoints: readonly (readonly number[])[];
    }[];
    readonly drawings: readonly {
      readonly id: string;
      readonly widthPt: number;
      readonly heightPt: number;
    }[];
  };
}

test('synthetic corpus preserves semantic geometry and real main/worker parity', async ({ page }) => {
  test.setTimeout(240_000);
  await page.goto('/tests/visual/conformance-fixture.html');
  await page.waitForFunction(() => document.body.dataset.status === 'ready');
  await expect.poll(() => page.evaluate(() =>
    Number(document.body.dataset.caseCount))).toBe(CONFORMANCE_CASES.length);

  for (let index = 0; index < CONFORMANCE_CASES.length; index += 1) {
    const testCase = CONFORMANCE_CASES[index]!;
    await test.step(testCase.id, async () => {
      const report = await page.evaluate(async (caseIndex) =>
        (window as unknown as {
          runConformanceCase: (index: number) => Promise<BrowserConformanceReport>;
        }).runConformanceCase(caseIndex), index);

      expect(report.caseId).toBe(testCase.id);
      expect(report.expected).toEqual(testCase.expected);
      expect(report.pageCount).toBe(testCase.expected.pageCount);
      expect(report.mainSizes).toEqual([{
        widthPt: testCase.expected.pageWidthPt,
        heightPt: testCase.expected.pageHeightPt,
      }]);
      expect(report.workerSizes).toEqual(report.mainSizes);
      expect(report.runsFinite).toBe(true);
      expect(report.targetTextPresent).toBe(true);
      expect(report.cloneSafe).toBe(true);
      expect(report.layoutFingerprint.length).toBeGreaterThan(0);
      expect(report.mainWorkerParity).toEqual({
        pageCount: true,
        sizes: true,
        runs: true,
      });

      // These values come entirely from authored twips/EMU and page margins.
      // Native Canvas shaping is deliberately absent from this exact
      // cross-browser comparison.
      expect(report.authoredGeometry.pageBoxes).toMatchObject([{
        widthPt: 612,
        heightPt: 792,
      }]);
      const pageBox = report.authoredGeometry.pageBoxes[0]!;
      expect(pageBox.contentTopPt).toBeGreaterThanOrEqual(0);
      expect(pageBox.contentBottomPt).toBeLessThanOrEqual(pageBox.heightPt);
      expect(pageBox.contentBottomPt).toBeGreaterThan(pageBox.contentTopPt);
      expect(report.authoredGeometry.tables).toHaveLength(testCase.expected.tableDepth);
      const tableWidths = report.authoredGeometry.tables
        .flatMap(({ columnWidthsPt }) => columnWidthsPt)
        .sort((left, right) => right - left);
      expect(tableWidths).toEqual(
        testCase.expected.tableDepth === 2
          ? [360, 180]
          : testCase.expected.tableDepth === 1
            ? [360]
            : [],
      );
      for (const table of report.authoredGeometry.tables) {
        expect(table.borderEndpoints.flat().every(Number.isFinite)).toBe(true);
      }
      expect(report.authoredGeometry.drawings).toHaveLength(
        testCase.expected.drawingCount,
      );
      for (const drawing of report.authoredGeometry.drawings) {
        expect(drawing).toMatchObject({ widthPt: 36, heightPt: 21.6 });
      }
    });
  }
});
