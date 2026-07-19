export interface ExactStateTracker {
  readonly states: string[];
  readonly seen: Set<string>;
}

export function createExactStateTracker(): ExactStateTracker {
  return { states: [], seen: new Set() };
}

export type ExactStateObservation =
  | Readonly<{ kind: 'advance' }>
  | Readonly<{ kind: 'fixed' | 'cycle'; states: readonly string[] }>;

export function observeExactState(
  tracker: ExactStateTracker,
  state: string,
): ExactStateObservation {
  const previous = tracker.states.at(-1);
  if (previous === state) {
    return { kind: 'fixed', states: Object.freeze([...tracker.states, state]) };
  }
  if (tracker.seen.has(state)) {
    return { kind: 'cycle', states: Object.freeze([...tracker.states, state]) };
  }
  tracker.states.push(state);
  tracker.seen.add(state);
  return { kind: 'advance' };
}
