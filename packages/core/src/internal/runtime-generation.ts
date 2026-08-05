type RuntimeState = 'uninitialized' | 'ready' | 'poisoned';

/** Environment-neutral realm initialization, generation, and poison fan-out. */
export class RuntimeGeneration<TFailure extends Error> {
  private state: RuntimeState = 'uninitialized';
  private generationValue = 0;
  private readiness: Promise<void> | undefined;
  private readonly poisonListeners = new Set<(error: TFailure) => void>();

  constructor(
    private readonly initialize: () => Promise<unknown> | unknown,
    private readonly reinitialize: () => Promise<unknown> | unknown,
    private readonly normalizeFailure: (error: unknown) => TFailure | null,
  ) {}

  get generation(): number { return this.generationValue; }
  get poisoned(): boolean { return this.state === 'poisoned'; }

  onPoison(listener: (error: TFailure) => void): () => void {
    this.poisonListeners.add(listener);
    return () => this.poisonListeners.delete(listener);
  }

  async ensureReady(): Promise<void> {
    if (this.state === 'ready') return;
    if (!this.readiness) {
      const operation = this.state === 'uninitialized' ? this.initialize : this.reinitialize;
      this.readiness = Promise.resolve().then(operation).then(() => {
        this.generationValue += 1;
        this.state = 'ready';
        this.readiness = undefined;
      }, (error: unknown) => {
        this.readiness = undefined;
        throw error;
      });
    }
    await this.readiness;
  }

  run<TResult>(operation: () => TResult): TResult {
    try {
      return operation();
    } catch (error) {
      const failure = this.normalizeFailure(error);
      if (!failure) throw error;
      this.poison(failure);
      throw failure;
    }
  }

  poison(error: TFailure): void {
    this.state = 'poisoned';
    this.readiness = undefined;
    for (const listener of this.poisonListeners) listener(error);
  }

  assertCurrent(generation: number): void {
    if (this.state !== 'ready' || generation !== this.generationValue) {
      throw new Error('WASM archive session belongs to a discarded runtime generation');
    }
  }
}
