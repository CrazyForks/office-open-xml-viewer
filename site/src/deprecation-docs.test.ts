import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { apiReference } from './lib/api-reference.js';

const page = readFileSync(new URL('./pages/deprecations.astro', import.meta.url), 'utf8');

describe('public deprecation documentation', () => {
  it('publishes the removal policy and every current public compatibility surface', () => {
    expect(page).toContain('scheduled for removal in a future breaking release');
    for (const api of [
      'maxZipEntryBytes',
      'OoxmlErrorSource',
      'XlsxChartSeries',
      'WasmParserHost.setWasmUrl()',
      'dropBitmapCacheByPath()',
    ]) {
      expect(page).toContain(api);
    }
  });

  it('links every documented maxZipEntryBytes option to its migration', () => {
    for (const classes of Object.values(apiReference)) {
      for (const apiClass of classes) {
        const option = apiClass.options?.find(({ name }) => name === 'maxZipEntryBytes');
        expect(option?.desc, apiClass.name).toContain('scheduled for removal in a future breaking release');
        expect(option?.detailsHref, apiClass.name).toBe('/deprecations#max-zip-entry-bytes');
      }
    }
  });
});
