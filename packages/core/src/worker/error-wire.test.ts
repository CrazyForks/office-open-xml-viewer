import { describe, expect, it } from 'vitest';
import {
  ParserResourceLimitError,
  type OoxmlErrorSource,
} from '../errors/ooxml-error';
import {
  deserializeWorkerError,
  parseResourceLimitError,
  serializeWorkerError,
} from './error-wire';

describe('worker error wire', () => {
  const rustError =
    'xlsx-parser error: OOXML_RESOURCE_LIMIT:{"code":"parser-resource-limit","stage":"decompression","source":"zip-part","part":"xl/worksheets/sheet1.xml","metric":"parsed-part-inflated-bytes","limit":5,"observed":6}';

  it('parses the Rust resource-limit envelope into an OoxmlError', () => {
    const error = parseResourceLimitError(new Error(rustError));
    expect(error).toBeInstanceOf(ParserResourceLimitError);
    expect(error).toMatchObject({
      code: 'parser-resource-limit',
      stage: 'decompression',
      source: 'zip-part',
      part: 'xl/worksheets/sheet1.xml',
      metric: 'parsed-part-inflated-bytes',
      limit: 5,
      observed: 6,
    });
  });

  it('survives worker serialization as a real typed error', () => {
    const error = deserializeWorkerError(serializeWorkerError(new Error(rustError)));
    expect(error).toBeInstanceOf(ParserResourceLimitError);
    expect(error).toMatchObject({
      code: 'parser-resource-limit',
      part: 'xl/worksheets/sheet1.xml',
    });
  });

  it('preserves unknown failures as ordinary errors', () => {
    const error = deserializeWorkerError(serializeWorkerError(new TypeError('bad input')));
    expect(error).toBeInstanceOf(Error);
    expect(error).not.toBeInstanceOf(ParserResourceLimitError);
    expect(error.name).toBe('TypeError');
    expect(error.message).toBe('bad input');
  });

  it('does not classify malformed envelopes as resource limits', () => {
    expect(parseResourceLimitError('OOXML_RESOURCE_LIMIT:{"code":"wrong"}')).toBeUndefined();
  });

  it('rejects unknown stage and source values at the worker boundary', () => {
    expect(
      parseResourceLimitError(
        new Error(
          'OOXML_RESOURCE_LIMIT:{"code":"parser-resource-limit","stage":"unknown","source":"zip-part","part":"word/document.xml","metric":"parsed-part-inflated-bytes","limit":5,"observed":6}',
        ),
      ),
    ).toBeUndefined();
    expect(
      deserializeWorkerError({
        message: 'invalid source',
        code: 'parser-resource-limit',
        stage: 'decompression',
        source: 'unknown' as OoxmlErrorSource,
        part: 'word/document.xml',
        metric: 'parsed-part-inflated-bytes',
        limit: 5,
        observed: 6,
      }),
    ).not.toBeInstanceOf(ParserResourceLimitError);
  });
});
