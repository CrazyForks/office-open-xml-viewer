import {
  OoxmlError,
  OoxmlResourceLimitError,
  type OoxmlFormat,
  type OoxmlResourceUsageSnapshot,
} from '../errors/ooxml-error.js';
import type {
  OoxmlResourceMetrics,
  OoxmlResourceMetricsCheckpoint,
} from '../types/resource-metrics.js';
import type { NormalizedOoxmlResourcePolicy } from './resource-policy.js';
import { emitOoxmlResourceDebugReport } from './resource-debug-view.js';
import { WasmTrapError } from './wasm-guard.js';

export interface OoxmlResourceMetricsSessionOptions {
  readonly enabled: boolean;
  readonly format: OoxmlFormat;
  readonly mode: OoxmlResourceMetrics['mode'];
  readonly scope?: OoxmlResourceMetrics['scope'];
  readonly policy: Readonly<NormalizedOoxmlResourcePolicy>;
  readonly now?: () => number;
  readonly onMetrics?: (report: OoxmlResourceMetrics) => void;
  /** Emit the human-readable console card in addition to `onMetrics`. */
  readonly emitToConsole?: boolean;
}

/**
 * Operation-scoped metrics collector shared by production observers and the
 * optional debug console renderer. It never records source addresses, ZIP
 * paths, error messages, document text, or passwords.
 */
export class OoxmlResourceMetricsSession {
  private readonly now: () => number;
  private readonly startedAt: number;
  private readonly policy: OoxmlResourceMetrics['policy'];
  private readonly checkpoints: OoxmlResourceMetricsCheckpoint[] = [];
  private sourceBytes?: number;
  private lastUsage?: OoxmlResourceUsageSnapshot;
  private mode: OoxmlResourceMetrics['mode'];
  private finished = false;

  constructor(private readonly options: OoxmlResourceMetricsSessionOptions) {
    this.now = options.now ?? defaultNow;
    this.startedAt = this.now();
    this.policy = Object.freeze({
      maxArchiveEntryBytes: options.policy.maxArchiveEntryBytes,
      maxTotalInflatedBytes: options.policy.maxTotalInflatedBytes,
    });
    this.mode = options.mode;
  }

  /** Record the realm that ultimately owns layout/rendering after capability
   * negotiation. This is intentionally mutable only until the report closes. */
  setMode(mode: OoxmlResourceMetrics['mode']): void {
    if (!this.options.enabled || this.finished) return;
    this.mode = mode;
  }

  setSourceBytes(bytes: number): void {
    if (!this.options.enabled) return;
    if (!Number.isSafeInteger(bytes) || bytes < 0) return;
    this.sourceBytes = bytes;
  }

  checkpoint(name: string, usage?: OoxmlResourceUsageSnapshot): void {
    if (!this.options.enabled || this.finished) return;
    if (usage) this.lastUsage = immutableUsage(usage);
    const checkpointUsage = this.lastUsage;
    this.checkpoints.push(Object.freeze({
      name: safeCheckpointName(name),
      elapsedMs: elapsed(this.startedAt, this.now()),
      ...(checkpointUsage ? { usage: checkpointUsage } : {}),
    }));
  }

  observeUsage(usage: OoxmlResourceUsageSnapshot | undefined): void {
    if (!this.options.enabled || this.finished || !usage) return;
    this.lastUsage = immutableUsage(usage);
  }

  succeed(outcome?: Readonly<Record<string, number>>): OoxmlResourceMetrics | undefined {
    return this.finish('ok', undefined, outcome);
  }

  fail(error: unknown): OoxmlResourceMetrics | undefined {
    return this.finish('error', error);
  }

  private finish(
    status: OoxmlResourceMetrics['status'],
    error?: unknown,
    outcome?: Readonly<Record<string, number>>,
  ): OoxmlResourceMetrics | undefined {
    if (!this.options.enabled || this.finished) return undefined;
    this.finished = true;
    const failureUsage = safeFailureUsage(error);
    // A typed violation is the authoritative terminal checkpoint. A previous
    // successful pull necessarily predates it and must not hide its counters.
    const finalUsage = failureUsage ?? this.lastUsage;
    const report: OoxmlResourceMetrics = Object.freeze({
      schemaVersion: 1,
      scope: this.options.scope ?? 'load',
      format: this.options.format,
      mode: this.mode,
      status,
      ...(this.sourceBytes === undefined ? {} : { sourceBytes: this.sourceBytes }),
      elapsedMs: elapsed(this.startedAt, this.now()),
      policy: this.policy,
      ...(finalUsage ? { usage: finalUsage } : {}),
      checkpoints: Object.freeze([...this.checkpoints]),
      ...(outcome ? { outcome: safeOutcome(outcome) } : {}),
      ...(status === 'error' ? { error: safeError(error) } : {}),
    });
    // Diagnostics must never change library semantics. A hostile/replaced
    // console or an application-provided observer cannot turn a successful
    // load into a failure or mask the original load error.
    if (this.options.onMetrics) safelyEmit(this.options.onMetrics, report);
    if (this.options.emitToConsole ?? this.options.onMetrics === undefined) {
      safelyEmit(emitOoxmlResourceDebugReport, report);
    }
    return report;
  }
}

function safelyEmit(
  emit: (report: OoxmlResourceMetrics) => void,
  report: OoxmlResourceMetrics,
): void {
  try {
    const result = emit(report) as unknown;
    if (isPromiseLike(result)) void Promise.resolve(result).catch(() => undefined);
  } catch {
    // Best-effort diagnostics and application telemetry only.
  }
}

function isPromiseLike(value: unknown): value is PromiseLike<unknown> {
  return (
    (typeof value === 'object' && value !== null) || typeof value === 'function'
  ) && typeof (value as { then?: unknown }).then === 'function';
}

function defaultNow(): number {
  return typeof performance === 'undefined' ? Date.now() : performance.now();
}

function immutableUsage(usage: OoxmlResourceUsageSnapshot): OoxmlResourceUsageSnapshot {
  return Object.freeze({
    archiveEntryCount: usage.archiveEntryCount,
    declaredInflatedBytes: usage.declaredInflatedBytes,
    ...(usage.largestInflatedEntryBytes === undefined
      ? {}
      : { largestInflatedEntryBytes: usage.largestInflatedEntryBytes }),
    distinctInflatedBytes: usage.distinctInflatedBytes,
    operationInflatedBytes: usage.operationInflatedBytes,
  });
}

function safeFailureUsage(error: unknown): OoxmlResourceUsageSnapshot | undefined {
  try {
    return error instanceof OoxmlResourceLimitError
      ? immutableUsage(error.details.violation.usage)
      : undefined;
  } catch {
    return undefined;
  }
}

function elapsed(start: number, end: number): number {
  return Math.max(0, Math.round((end - start) * 10) / 10);
}

function safeCheckpointName(value: string): string {
  const cleaned = value.replace(/[^a-z0-9 -]/giu, '').trim().slice(0, 32);
  return cleaned || 'checkpoint';
}

function safeOutcome(value: Readonly<Record<string, number>>): Readonly<Record<string, number>> {
  return Object.freeze(Object.fromEntries(Object.entries(value).filter(
    ([key, count]) => /^[a-z][a-z0-9-]{0,31}$/u.test(key)
      && Number.isSafeInteger(count) && count >= 0,
  )));
}

function safeError(error: unknown): OoxmlResourceMetrics['error'] {
  try {
    if (error instanceof OoxmlResourceLimitError) {
      const violation = error.details.violation;
      return Object.freeze({
        code: error.code,
        stage: error.details.stage,
        resource: violation.resource,
        metric: violation.metric,
      });
    }
    if (error instanceof OoxmlError) {
      return Object.freeze({ code: error.code });
    }
    if (
      error instanceof WasmTrapError
      || (error !== null
        && typeof error === 'object'
        && (error as { code?: unknown }).code === 'parser-crashed')
    ) {
      return Object.freeze({ code: 'parser-crashed' });
    }
  } catch {
    // Proxies and hostile accessors are treated as unclassified errors.
  }
  return Object.freeze({});
}
