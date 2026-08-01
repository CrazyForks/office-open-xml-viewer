import { describe, it, expect } from 'vitest';
import {
  OoxmlError,
  OoxmlResourceLimitError,
  type OoxmlErrorCode,
  type OoxmlResourceMetric,
  type OoxmlResourceName,
} from './ooxml-error';

describe('OoxmlError', () => {
  it('is an Error subclass carrying a machine-readable code', () => {
    const err = new OoxmlError('encrypted', 'This file is password-protected.');
    expect(err).toBeInstanceOf(Error);
    expect(err).toBeInstanceOf(OoxmlError);
    expect(err.code).toBe('encrypted');
    expect(err.message).toBe('This file is password-protected.');
  });

  it('sets name to "OoxmlError" so stringification is stable', () => {
    const err = new OoxmlError('not-ooxml', 'nope');
    expect(err.name).toBe('OoxmlError');
    // Error.prototype.toString uses `name: message`.
    expect(String(err)).toBe('OoxmlError: nope');
  });

  it('keeps the code readonly at the type level and reflects each variant', () => {
    const codes: OoxmlErrorCode[] = [
      'encrypted',
      'legacy-binary-format',
      'not-ooxml',
    ];
    for (const code of codes) {
      expect(new OoxmlError(code, code).code).toBe(code);
    }
  });

  it('uses a separate additive error class with structured resource details', () => {
    const err = new OoxmlResourceLimitError('too large', {
      stage: 'decompression',
      violation: {
        format: 'docx',
        operation: 'parse',
        resource: 'archive-entry',
        part: 'word/document.xml',
        metric: 'actual-inflated-bytes',
        limit: 10,
        observed: 11,
        configurable: true,
        usage: {
          archiveEntryCount: 1,
          declaredInflatedBytes: 11,
          largestInflatedEntryBytes: 11,
          distinctInflatedBytes: 11,
          operationInflatedBytes: 11,
        },
      },
    });
    expect(err).toBeInstanceOf(Error);
    expect(err).toBeInstanceOf(OoxmlResourceLimitError);
    expect(err.code).toBe('ooxml-resource-limit');
    expect(err.details.violation).toMatchObject({
      format: 'docx',
      operation: 'parse',
      resource: 'archive-entry',
      part: 'word/document.xml',
      metric: 'actual-inflated-bytes',
      usage: { largestInflatedEntryBytes: 11 },
    });
    expect(Object.isFrozen(err.details)).toBe(true);
    expect(Object.isFrozen(err.details.violation)).toBe(true);
    expect(Object.isFrozen(err.details.violation.usage)).toBe(true);
  });

  it('keeps future resource and metric literals source-compatible', () => {
    const resource: OoxmlResourceName = 'slide-wire';
    const metric: OoxmlResourceMetric = 'retained-bytes';
    expect({ resource, metric }).toEqual({
      resource: 'slide-wire',
      metric: 'retained-bytes',
    });
  });

  it('captures a stack trace', () => {
    const err = new OoxmlError('legacy-binary-format', 'legacy');
    expect(typeof err.stack).toBe('string');
  });
});
