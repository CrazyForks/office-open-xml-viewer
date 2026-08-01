import { OoxmlResourceLimitError, type OoxmlFormat, type OoxmlResourceUsageSnapshot } from '../errors/ooxml-error.js';
import type { NormalizedOoxmlResourcePolicy } from './resource-policy.js';
import { emitOoxmlResourceDebugReport } from './resource-debug-view.js';

export interface OoxmlResourceDebugCheckpoint {
  readonly name: string;
  readonly elapsedMs: number;
  readonly usage?: OoxmlResourceUsageSnapshot;
}

export interface OoxmlResourceDebugReport {
  readonly format: OoxmlFormat;
  readonly mode: 'main' | 'worker' | 'node';
  readonly status: 'ok' | 'error';
  readonly sourceBytes?: number;
  readonly elapsedMs: number;
  readonly policy: Readonly<NormalizedOoxmlResourcePolicy>;
  readonly usage?: OoxmlResourceUsageSnapshot;
  readonly checkpoints: readonly OoxmlResourceDebugCheckpoint[];
  readonly outcome?: Readonly<Record<string, number>>;
  readonly error?: Readonly<{
    code?: string;
    stage?: string;
    resource?: string;
    metric?: string;
  }>;
}

export interface OoxmlResourceDebugSessionOptions {
  readonly enabled: boolean;
  readonly format: OoxmlFormat;
  readonly mode: OoxmlResourceDebugReport['mode'];
  readonly policy: Readonly<NormalizedOoxmlResourcePolicy>;
  readonly now?: () => number;
  readonly emit?: (report: OoxmlResourceDebugReport) => void;
}

/**
 * Load-scoped diagnostic collector. It never records source addresses, ZIP
 * paths, error messages, document text, or passwords; callers may safely leave
 * it enabled while investigating admission limits.
 */
export class OoxmlResourceDebugSession {
  private readonly now: () => number;
  private readonly startedAt: number;
  private readonly checkpoints: OoxmlResourceDebugCheckpoint[] = [];
  private sourceBytes?: number;
  private lastUsage?: OoxmlResourceUsageSnapshot;
  private finished = false;

  constructor(private readonly options: OoxmlResourceDebugSessionOptions) {
    this.now = options.now ?? defaultNow;
    this.startedAt = this.now();
  }

  setSourceBytes(bytes: number): void {
    if (!this.options.enabled) return;
    if (!Number.isSafeInteger(bytes) || bytes < 0) return;
    this.sourceBytes = bytes;
  }

  checkpoint(name: string, usage?: OoxmlResourceUsageSnapshot): void {
    if (!this.options.enabled || this.finished) return;
    if (usage) this.lastUsage = usage;
    const checkpointUsage = usage ?? this.lastUsage;
    this.checkpoints.push(Object.freeze({
      name: safeCheckpointName(name),
      elapsedMs: elapsed(this.startedAt, this.now()),
      ...(checkpointUsage ? { usage: checkpointUsage } : {}),
    }));
  }

  observeUsage(usage: OoxmlResourceUsageSnapshot | undefined): void {
    if (!this.options.enabled || this.finished || !usage) return;
    this.lastUsage = usage;
  }

  succeed(outcome?: Readonly<Record<string, number>>): OoxmlResourceDebugReport | undefined {
    return this.finish('ok', undefined, outcome);
  }

  fail(error: unknown): OoxmlResourceDebugReport | undefined {
    return this.finish('error', error);
  }

  private finish(
    status: OoxmlResourceDebugReport['status'],
    error?: unknown,
    outcome?: Readonly<Record<string, number>>,
  ): OoxmlResourceDebugReport | undefined {
    if (!this.options.enabled || this.finished) return undefined;
    this.finished = true;
    const failureUsage = error instanceof OoxmlResourceLimitError
      ? error.details.violation.usage
      : undefined;
    const finalUsage = this.lastUsage ?? failureUsage;
    const report: OoxmlResourceDebugReport = Object.freeze({
      format: this.options.format,
      mode: this.options.mode,
      status,
      ...(this.sourceBytes === undefined ? {} : { sourceBytes: this.sourceBytes }),
      elapsedMs: elapsed(this.startedAt, this.now()),
      policy: this.options.policy,
      ...(finalUsage ? { usage: finalUsage } : {}),
      checkpoints: Object.freeze([...this.checkpoints]),
      ...(outcome ? { outcome: safeOutcome(outcome) } : {}),
      ...(status === 'error' ? { error: safeError(error) } : {}),
    });
    (this.options.emit ?? emitOoxmlResourceDebugReport)(report);
    return report;
  }
}

function defaultNow(): number {
  return typeof performance === 'undefined' ? Date.now() : performance.now();
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

function safeError(error: unknown): OoxmlResourceDebugReport['error'] {
  if (error instanceof OoxmlResourceLimitError) {
    const violation = error.details.violation;
    return Object.freeze({
      code: error.code,
      stage: error.details.stage,
      resource: violation.resource,
      metric: violation.metric,
    });
  }
  if (!error || typeof error !== 'object') return Object.freeze({});
  const value = error as { code?: unknown; stage?: unknown };
  return Object.freeze({
    ...(safeIdentifier(value.code) ? { code: value.code } : {}),
    ...(safeIdentifier(value.stage) ? { stage: value.stage } : {}),
  });
}

function safeIdentifier(value: unknown): value is string {
  return typeof value === 'string' && /^[a-z0-9][a-z0-9-]{0,63}$/u.test(value);
}
