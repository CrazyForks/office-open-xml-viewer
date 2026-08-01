import { describe, expect, it, vi } from 'vitest';
import { OoxmlResourceLimitError } from '../errors/ooxml-error.js';
import {
  OoxmlResourceDebugSession,
} from './resource-debug.js';
import { formatOoxmlResourceDebugReport } from './resource-debug-view.js';

const policy = {
  maxArchiveEntryBytes: 128 * 1024 * 1024,
  maxTotalInflatedBytes: 256 * 1024 * 1024,
} as const;

const usage = {
  archiveEntryCount: 42,
  declaredInflatedBytes: 12 * 1024 * 1024,
  largestInflatedEntryBytes: 32 * 1024 * 1024,
  distinctInflatedBytes: 80 * 1024 * 1024,
  operationInflatedBytes: 6 * 1024 * 1024,
} as const;

describe('OoxmlResourceDebugSession', () => {
  it('emits one data-safe success report with limit-setting metrics', () => {
    const emit = vi.fn();
    const ticks = [0, 12, 30, 45];
    const session = new OoxmlResourceDebugSession({
      enabled: true,
      format: 'docx',
      mode: 'worker',
      policy,
      now: () => ticks.shift() ?? 45,
      emit,
    });
    session.setSourceBytes(4 * 1024 * 1024);
    session.checkpoint('container ready');
    session.checkpoint('model streamed', usage);
    const report = session.succeed({ pages: 12 });
    session.fail(new Error('must not be emitted'));

    expect(emit).toHaveBeenCalledTimes(1);
    expect(report).toMatchObject({
      format: 'docx',
      mode: 'worker',
      status: 'ok',
      sourceBytes: 4 * 1024 * 1024,
      elapsedMs: 45,
      usage,
      outcome: { pages: 12 },
    });
    expect(formatOoxmlResourceDebugReport(report!)).toContain(
      'largest entry  ████░░░░░░░░░░░░  32.0 MiB / 128 MiB',
    );
  });

  it('reports only stable failure discriminants, never error text or part paths', () => {
    const emit = vi.fn();
    const session = new OoxmlResourceDebugSession({
      enabled: true,
      format: 'xlsx',
      mode: 'main',
      policy,
      now: () => 0,
      emit,
    });
    const error = new OoxmlResourceLimitError('secret source and cell text', {
      stage: 'decompression',
      violation: {
        format: 'xlsx',
        operation: 'parse',
        resource: 'archive-entry',
        metric: 'actual-inflated-bytes',
        part: 'xl/media/private.png',
        limit: 10,
        observed: 11,
        configurable: true,
        usage,
      },
    });
    const report = session.fail(error)!;
    const text = formatOoxmlResourceDebugReport(report);

    expect(report.error).toEqual({
      code: 'ooxml-resource-limit',
      stage: 'decompression',
      resource: 'archive-entry',
      metric: 'actual-inflated-bytes',
    });
    expect(report.usage).toEqual(usage);
    expect(text).not.toContain('secret');
    expect(text).not.toContain('private.png');
  });

  it('does no work or output when disabled', () => {
    const emit = vi.fn();
    const now = vi.fn(() => 0);
    const session = new OoxmlResourceDebugSession({
      enabled: false,
      format: 'pptx',
      mode: 'main',
      policy,
      now,
      emit,
    });
    session.checkpoint('ignored', usage);
    expect(session.succeed()).toBeUndefined();
    expect(emit).not.toHaveBeenCalled();
    expect(now).toHaveBeenCalledTimes(1);
  });
});
