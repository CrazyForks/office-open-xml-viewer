import { describe, expect, it } from 'vitest';
import { apiReference } from './api-reference.js';

describe('official-site API reference', () => {
  it('documents the shared resource controls on every browser API class', () => {
    for (const classes of Object.values(apiReference)) {
      for (const apiClass of classes) {
        const optionNames = apiClass.options?.map(({ name }) => name) ?? [];
        expect(optionNames, apiClass.name).toEqual(expect.arrayContaining([
          'resourceLimits',
          'onResourceMetrics',
          'debug',
        ]));
      }
    }
  });

  it('keeps every semantic emphasis synchronized with its description', () => {
    for (const classes of Object.values(apiReference)) {
      for (const apiClass of classes) {
        for (const item of [...(apiClass.options ?? []), ...apiClass.methods]) {
          if (item.emphasis) {
            expect(item.desc, `${apiClass.name}: ${'name' in item ? item.name : item.sig}`)
              .toContain(item.emphasis);
          }
        }
      }
    }
  });

  it('documents the Viewer error-delivery contract and typed resource failures', () => {
    for (const classes of Object.values(apiReference)) {
      for (const apiClass of classes.filter(({ name }) => name.endsWith('Viewer'))) {
        const onError = apiClass.options?.find(({ name }) => name === 'onError');
        expect(onError, apiClass.name).toBeDefined();
        expect(onError?.desc, apiClass.name).toContain('load() resolves; without it load() rejects');
        expect(onError?.desc, apiClass.name).toContain('OoxmlResourceLimitError');
        expect(onError?.desc, apiClass.name).toContain('OoxmlDecodedImageLimitError');
        expect(onError?.desc, apiClass.name).toContain('message text is not a stable discriminator');
      }
    }
  });
});
