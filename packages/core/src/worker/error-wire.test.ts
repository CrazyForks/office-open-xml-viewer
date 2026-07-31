import { describe, expect, it } from 'vitest';
import { OoxmlResourceLimitError } from '../errors/ooxml-error.js';
import {
  deserializeWorkerError,
  parseResourceLimitError,
  serializeWorkerError,
} from './error-wire.js';

const USAGE = {
  archiveEntryCount: 1,
  declaredInflatedBytes: 6,
  distinctInflatedBytes: 6,
  operationInflatedBytes: 6,
};

describe('worker error wire', () => {
  const rustError =
    'OOXML_RESOURCE_LIMIT:{"code":"ooxml-resource-limit","details":{"stage":"decompression","violation":{"format":"xlsx","operation":"parse","resource":"archive-entry","part":"xl/worksheets/sheet1.xml","metric":"actual-inflated-bytes","limit":5,"observed":6,"configurable":true,"usage":{"archiveEntryCount":1,"declaredInflatedBytes":6,"distinctInflatedBytes":6,"operationInflatedBytes":6}}}}';

  it('parses an exact Rust envelope into the public typed error', () => {
    const error = parseResourceLimitError(new Error(rustError));
    expect(error).toBeInstanceOf(OoxmlResourceLimitError);
    expect(error).toMatchObject({ code: 'ooxml-resource-limit' });
    expect(error?.details.violation).toMatchObject({
      format: 'xlsx',
      operation: 'parse',
      resource: 'archive-entry',
      part: 'xl/worksheets/sheet1.xml',
      metric: 'actual-inflated-bytes',
      limit: 5,
      observed: 6,
      configurable: true,
    });
  });

  it('survives worker serialization and structured clone as a real typed error', () => {
    const wire = structuredClone(serializeWorkerError(new Error(rustError)));
    const error = deserializeWorkerError(wire);
    expect(error).toBeInstanceOf(OoxmlResourceLimitError);
    expect((error as OoxmlResourceLimitError).details.violation).toMatchObject({
      resource: 'archive-entry',
      part: 'xl/worksheets/sheet1.xml',
    });
  });

  it('represents an archive-wide metric without a dummy part', () => {
    const error = new OoxmlResourceLimitError('too many entries', {
      stage: 'container',
      violation: {
        format: 'docx',
        operation: 'open',
        resource: 'archive',
        metric: 'entry-count',
        limit: 20_000,
        observed: 20_001,
        configurable: false,
        usage: { ...USAGE, archiveEntryCount: 20_001 },
      },
    });
    const restored = deserializeWorkerError(serializeWorkerError(error));
    expect(restored).toBeInstanceOf(OoxmlResourceLimitError);
    expect((restored as OoxmlResourceLimitError).details.violation).not.toHaveProperty('part');
  });

  it('preserves the hard central-directory metadata limit', () => {
    const error = new OoxmlResourceLimitError('central directory is too large', {
      stage: 'container',
      violation: {
        format: 'xlsx',
        operation: 'open',
        resource: 'archive',
        metric: 'central-directory-bytes',
        limit: 16 * 1024 * 1024,
        observed: 16 * 1024 * 1024 + 1,
        configurable: false,
        usage: USAGE,
      },
    });
    const restored = deserializeWorkerError(serializeWorkerError(error));
    expect(restored).toBeInstanceOf(OoxmlResourceLimitError);
    expect((restored as OoxmlResourceLimitError).details.violation).toMatchObject({
      resource: 'archive',
      metric: 'central-directory-bytes',
      configurable: false,
    });
    expect((restored as OoxmlResourceLimitError).details.violation).not.toHaveProperty('part');
  });

  it('preserves non-configurable parser-buffer limits and their part context', () => {
    const error = new OoxmlResourceLimitError('worksheet row is too large', {
      stage: 'parsing',
      violation: {
        format: 'xlsx',
        operation: 'parse-sheet',
        resource: 'worksheet-row',
        metric: 'projected-bytes',
        part: 'xl/worksheets/sheet1.xml',
        limit: 8 * 1024 * 1024,
        observed: 8 * 1024 * 1024 + 1,
        configurable: false,
        usage: USAGE,
      },
    });
    const restored = deserializeWorkerError(serializeWorkerError(error));
    expect(restored).toBeInstanceOf(OoxmlResourceLimitError);
    expect((restored as OoxmlResourceLimitError).details).toMatchObject({
      stage: 'parsing',
      violation: {
        resource: 'worksheet-row',
        metric: 'projected-bytes',
        part: 'xl/worksheets/sheet1.xml',
        configurable: false,
      },
    });
  });

  it('preserves a non-configurable hard per-entry violation', () => {
    const hardError = rustError.replace('"configurable":true', '"configurable":false');
    const error = parseResourceLimitError(hardError);
    expect(error).toBeInstanceOf(OoxmlResourceLimitError);
    expect(error?.details.violation.configurable).toBe(false);
  });

  it.each([new TypeError('bad input'), new RangeError('bad range')])(
    'preserves %s across the worker wire',
    (original) => {
      const error = deserializeWorkerError(serializeWorkerError(original));
      expect(error).toBeInstanceOf(original.constructor);
      expect(error.message).toBe(original.message);
    },
  );

  it('does not classify malformed or wrapped envelopes as resource limits', () => {
    expect(parseResourceLimitError('OOXML_RESOURCE_LIMIT:{"code":"wrong"}')).toBeUndefined();
    expect(parseResourceLimitError(`xlsx parser: ${rustError}`)).toBeUndefined();
  });

  it('rejects invalid discriminants at the worker boundary', () => {
    expect(
      deserializeWorkerError({
        message: 'invalid details',
        code: 'ooxml-resource-limit',
        resourceLimit: {
          stage: 'container',
          violation: {
            format: 'pptx',
            operation: 'open',
            resource: 'archive-entry',
            metric: 'actual-inflated-bytes',
            part: 'ppt/slides/slide1.xml',
            limit: 5,
            observed: 6,
            configurable: true,
            usage: USAGE,
          },
        },
      }),
    ).not.toBeInstanceOf(OoxmlResourceLimitError);

    const wrongParserStage = new OoxmlResourceLimitError('wrong stage', {
      stage: 'parsing',
      violation: {
        format: 'xlsx',
        operation: 'parse-sheet',
        resource: 'xml-event',
        metric: 'bytes',
        limit: 1,
        observed: 2,
        configurable: false,
        usage: USAGE,
      },
    });
    const payload = structuredClone(serializeWorkerError(wrongParserStage));
    if (payload.resourceLimit) Object.assign(payload.resourceLimit, { stage: 'container' });
    expect(deserializeWorkerError(payload)).not.toBeInstanceOf(OoxmlResourceLimitError);
  });

  it('rejects removed declared-total metrics and fractional wire counters', () => {
    const declaredTotal = structuredClone(serializeWorkerError(new Error(rustError)));
    if (declaredTotal.resourceLimit) {
      Object.assign(declaredTotal.resourceLimit.violation, {
        resource: 'archive',
        metric: 'declared-total-inflated-bytes',
        configurable: true,
      });
      delete (declaredTotal.resourceLimit.violation as { part?: string }).part;
    }
    expect(deserializeWorkerError(declaredTotal)).not.toBeInstanceOf(OoxmlResourceLimitError);

    const fractional = structuredClone(serializeWorkerError(new Error(rustError)));
    if (fractional.resourceLimit) fractional.resourceLimit.violation.limit = 5.5;
    expect(deserializeWorkerError(fractional)).not.toBeInstanceOf(OoxmlResourceLimitError);
  });
});
