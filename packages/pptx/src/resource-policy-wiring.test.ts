import { describe, expect, it, vi } from 'vitest';
import { PptxPresentation } from './presentation.js';

describe('PptxPresentation resource-policy wiring', () => {
  it('forwards one normalized policy object to worker parsing', async () => {
    let request: Record<string, unknown> | undefined;
    const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
    instance._mode = 'worker';
    instance._bridge = {
      request: vi.fn(async (createRequest: (id: number) => Record<string, unknown>) => {
        request = createRequest(3);
        return {
          kind: 'parsedMeta',
          id: 3,
          meta: {
            slideCount: 0,
            slideWidth: 0,
            slideHeight: 0,
            majorFont: null,
            minorFont: null,
            notes: [],
            mediaElements: [],
            hidden: [],
            partNames: [],
          },
        };
      }),
    };
    const policy = {
      maxArchiveEntryBytes: null,
      maxTotalInflatedBytes: 512,
    } as const;

    await (
      instance as unknown as {
        _parse(
          data: ArrayBuffer,
          resourcePolicy: typeof policy,
          useGoogleFonts: boolean,
          timeout: number,
        ): Promise<void>;
      }
    )._parse(new ArrayBuffer(1), policy, false, 30_000);

    expect(request).toMatchObject({ kind: 'parse', id: 3, resourcePolicy: policy });
    expect(request).not.toHaveProperty('maxZipEntryBytes');
    expect(request).not.toHaveProperty('parserResourceLimits');
  });
});
