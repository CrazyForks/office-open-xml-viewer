import type { DeepReadonly, DocumentLayout, LayoutServices } from './types.js';
import type { DocRun } from '../types.js';
import { mathResourceKey, type MathOccurrence } from './resources.js';

export interface DocumentLayoutRuntimeState {
  services: LayoutServices | null;
  retainedErrorLayout: DeepReadonly<DocumentLayout> | null;
  readonly defaultCurrentDateMs: number;
}

const documentLayoutRuntime = Symbol('document-layout-runtime');

type RuntimeOwner = object & {
  [documentLayoutRuntime]?: DocumentLayoutRuntimeState;
};

export function attachDocumentLayoutRuntime(
  owner: object,
  defaultCurrentDateMs: number,
): void {
  Object.defineProperty(owner, documentLayoutRuntime, {
    configurable: false,
    enumerable: false,
    writable: false,
    value: { services: null, retainedErrorLayout: null, defaultCurrentDateMs },
  });
}

export function documentLayoutRuntimeOf(owner: object): DocumentLayoutRuntimeState {
  const runtime = (owner as RuntimeOwner)[documentLayoutRuntime];
  if (runtime) return runtime;
  throw new Error('Document layout runtime is not initialized; attach it explicitly');
}

export interface ImmutableResourceLookup<T> {
  readonly keys: readonly string[];
  resolve(resourceKey: string): T;
}

/** Keep browser/DOM handles in a private closure while exposing fixed immutable membership. */
export function createImmutableResourceLookup<T>(entries: ReadonlyMap<string, T>): ImmutableResourceLookup<T> {
  const snapshot = new Map(entries);
  const keys = Object.freeze([...snapshot.keys()].sort());
  return Object.freeze({
    keys,
    resolve(resourceKey: string): T {
      const value = snapshot.get(resourceKey);
      if (value === undefined) throw new Error(`Unknown runtime resource: ${resourceKey}`);
      return value;
    },
  });
}

const privateResourceLookups = new WeakMap<object, ImmutableResourceLookup<unknown>>();

export function attachPrivateResourceLookup<T>(
  owner: object,
  entries: ReadonlyMap<string, T>,
  expectedKeys: Iterable<string> = entries.keys(),
): void {
  if (privateResourceLookups.has(owner)) {
    throw new Error('Private resource lookup is already attached');
  }
  const actual = new Set(entries.keys());
  const expected = new Set(expectedKeys);
  const missing = [...expected].filter((key) => !actual.has(key)).sort();
  const extra = [...actual].filter((key) => !expected.has(key)).sort();
  if (missing.length > 0 || extra.length > 0) {
    throw new Error(
      `Runtime resource membership mismatch: missing [${missing.join(', ')}]; extra [${extra.join(', ')}]`,
    );
  }
  privateResourceLookups.set(owner, createImmutableResourceLookup(entries));
}

export function privateResourceLookupOf<T>(owner: object): ImmutableResourceLookup<T> | undefined {
  return privateResourceLookups.get(owner) as ImmutableResourceLookup<T> | undefined;
}

export interface MathOccurrenceLookup {
  resourceKey(runs: readonly DocRun[], runIndex: number): string;
}

const mathOccurrenceLookups = new WeakMap<object, MathOccurrenceLookup>();

/** Address parser runs by their owning array and ordinal. Formula contents are
 * deliberately absent from this private registry and every public snapshot. */
export function attachMathOccurrenceLookup(
  owner: object,
  occurrences: readonly MathOccurrence[],
): void {
  if (mathOccurrenceLookups.has(owner)) throw new Error('Math occurrence lookup is already attached');
  const byRuns = new WeakMap<readonly DocRun[], ReadonlyMap<number, string>>();
  const mutable = new Map<readonly DocRun[], Map<number, string>>();
  for (const occurrence of occurrences) {
    const entries = mutable.get(occurrence.runs) ?? new Map<number, string>();
    if (entries.has(occurrence.runIndex)) throw new Error('Duplicate math run occurrence');
    entries.set(
      occurrence.runIndex,
      mathResourceKey(occurrence.source, occurrence.display ? 'display' : 'inline'),
    );
    mutable.set(occurrence.runs, entries);
  }
  for (const [runs, entries] of mutable) byRuns.set(runs, new Map(entries));
  mathOccurrenceLookups.set(owner, Object.freeze({
    resourceKey(runs: readonly DocRun[], runIndex: number): string {
      const key = byRuns.get(runs)?.get(runIndex);
      if (!key) throw new Error(`Unknown math occurrence at run ${runIndex}`);
      return key;
    },
  }));
}

export function mathOccurrenceLookupOf(owner: object): MathOccurrenceLookup | undefined {
  return mathOccurrenceLookups.get(owner);
}
