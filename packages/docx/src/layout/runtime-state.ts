import type { LayoutServices } from './types.js';
import type { PaintResourceRegistry } from './types.js';
import type { NumberFormat } from '@silurus/ooxml-core';
import type { BodyLayoutKernel } from './body-layout-kernel.js';
import type { LayoutVariantStore } from './variant-store.js';

export interface DocumentLayoutRuntimeState {
  services: LayoutServices | null;
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
    value: { services: null, defaultCurrentDateMs },
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
type LayoutServicesRuntimeOwner = object;

const layoutServicesRuntimeOwners = new WeakMap<object, LayoutServicesRuntimeOwner>();
const bodyLayoutKernels = new WeakMap<LayoutServicesRuntimeOwner, BodyLayoutKernel>();
const layoutVariantStores = new WeakMap<LayoutServices, LayoutVariantStore>();

/** Service views may replace one component, but mixing components already owned
 * by different documents must fail before acquisition can use a foreign kernel. */
function layoutServicesRuntimeOwner(
  services: LayoutServices,
  create: boolean,
): LayoutServicesRuntimeOwner | undefined {
  const components = [services.text, services.images, services.math];
  const owners = new Set(components.flatMap((component) => {
    const owner = layoutServicesRuntimeOwners.get(component);
    return owner ? [owner] : [];
  }));
  if (owners.size > 1) {
    throw new Error('Layout services combine foreign runtime owners');
  }
  const owner = owners.values().next().value as LayoutServicesRuntimeOwner | undefined;
  const unowned = components.filter((component) => !layoutServicesRuntimeOwners.has(component));
  if (owner && unowned.length > 1) {
    throw new Error('Layout services are missing service lineage for multiple components');
  }
  if (!owner && !create) return undefined;
  const resolved = owner ?? {};
  for (const component of components) {
    const existing = layoutServicesRuntimeOwners.get(component);
    if (existing && existing !== resolved) {
      throw new Error('Layout services combine foreign runtime owners');
    }
    layoutServicesRuntimeOwners.set(component, resolved);
  }
  return resolved;
}

export function attachBodyLayoutKernel(
  services: LayoutServices,
  kernel: BodyLayoutKernel,
): void {
  const owner = layoutServicesRuntimeOwner(services, true)!;
  if (bodyLayoutKernels.has(owner)) throw new Error('Body layout kernel is already attached');
  bodyLayoutKernels.set(owner, kernel);
}

export function bodyLayoutKernelOf(services: LayoutServices): BodyLayoutKernel | undefined {
  const owner = layoutServicesRuntimeOwner(services, false);
  return owner ? bodyLayoutKernels.get(owner) : undefined;
}

export function attachLayoutVariantStore(
  services: LayoutServices,
  store: LayoutVariantStore,
): void {
  if (layoutVariantStores.has(services)) throw new Error('Layout variant store is already attached');
  layoutVariantStores.set(services, store);
}

export function layoutVariantStoreOf(services: LayoutServices): LayoutVariantStore | undefined {
  return layoutVariantStores.get(services);
}

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

const paintResourceRegistries = new WeakMap<object, PaintResourceRegistry>();

export interface FieldAcquisitionContext {
  readonly totalPages: number;
  /** Resolve one PAGE field occurrence from the preceding pagination iteration. */
  readonly resolvePageField?: (
    paragraph: object,
    sourceRunIndex: number,
  ) => PageFieldAcquisitionContext | undefined;
  /** Resolve a generated table-row occurrence from the preceding iteration. */
  readonly resolveTablePageField?: (
    occurrenceId: string,
    paragraph: object,
    sourceRunIndex: number,
  ) => PageFieldAcquisitionContext | undefined;
  /** Page numbering is finalized only after one pagination iteration. */
  readonly resolveDestinationPage?: (
    physicalPageIndex: number,
  ) => PageFieldAcquisitionContext | undefined;
  readonly resolveTableOccurrencePage?: (
    occurrenceId: string,
  ) => PageFieldAcquisitionContext | undefined;
}

export interface PageFieldAcquisitionContext {
  readonly pageIndex: number;
  readonly displayPageNumber: number;
  readonly pageNumberFormat: NumberFormat;
}

const fieldAcquisitionContexts = new WeakMap<object, FieldAcquisitionContext>();

export function createLayoutServicesRuntimeView(
  services: LayoutServices,
  overrides: Readonly<{ text?: LayoutServices['text'] }> = {},
): LayoutServices {
  const view = Object.freeze({ ...services, ...overrides });
  const kernel = bodyLayoutKernelOf(services);
  if (!kernel) throw new Error('Body layout kernel is not attached to the supplied services');
  if (bodyLayoutKernelOf(view) !== kernel) {
    throw new Error('Layout service view did not retain its body layout kernel owner');
  }
  const lookup = privateResourceLookups.get(services);
  if (lookup) privateResourceLookups.set(view, lookup);
  const registry = paintResourceRegistries.get(services);
  if (registry) paintResourceRegistries.set(view, registry);
  return view;
}

/** Create one immutable service identity for a pagination-field iteration. */
export function createFieldAcquisitionServicesView(
  services: LayoutServices,
  context: FieldAcquisitionContext,
): LayoutServices {
  if (!Number.isInteger(context.totalPages) || context.totalPages < 1) {
    throw new RangeError('Field acquisition totalPages must be a positive integer');
  }
  const view = createLayoutServicesRuntimeView(services);
  fieldAcquisitionContexts.set(view, Object.freeze({ ...context }));
  return view;
}

export function fieldAcquisitionContextOf(owner: object): FieldAcquisitionContext {
  return fieldAcquisitionContexts.get(owner) ?? Object.freeze({ totalPages: 1 });
}

/** Associate cloneable resource descriptors with their document-scoped owner
 * without widening the stable LayoutServices or public document contracts. */
export function attachPaintResourceRegistry(
  owner: object,
  registry: PaintResourceRegistry,
): void {
  if (paintResourceRegistries.has(owner)) {
    throw new Error('Paint resource registry is already attached');
  }
  paintResourceRegistries.set(owner, registry);
}

export function paintResourceRegistryOf(owner: object): PaintResourceRegistry {
  const registry = paintResourceRegistries.get(owner);
  if (!registry) throw new Error('Paint resource registry is not attached');
  return registry;
}
