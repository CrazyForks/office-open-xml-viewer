import type { WorkerBridgeTransport } from '@silurus/ooxml-core';
import {
  BoundedPullSession,
  HARD_MAX_PPTX_SLIDE_JSON_BYTES,
  PULL_SESSION_INSUFFICIENT_CREDIT_CODE,
  PULL_SESSION_PROTOCOL,
  type PullSessionIdentity,
  type PullSessionResponse,
} from '@silurus/ooxml-core/worker';
import type { Slide } from './types.js';

export const PPTX_INITIAL_SLIDE_PULL_BYTES = 1024 * 1024;

export interface PptxSlidePullClientOptions {
  readonly slideCount: number;
  readonly generation?: number;
  readonly transport: WorkerBridgeTransport<PullSessionResponse<ArrayBuffer, number>>;
  readonly open: (
    slideIndex: number,
    identity: PullSessionIdentity<number>,
    timeoutMs?: number,
  ) => Promise<void>;
}

/** Main-realm owner of PPTX's one-slide transferred pull sessions. */
export class PptxSlidePullClient {
  private readonly active = new Set<BoundedPullSession<ArrayBuffer, number>>();
  private nextSessionId = 1;

  constructor(private readonly options: PptxSlidePullClientOptions) {
    if (!Number.isSafeInteger(options.slideCount) || options.slideCount < 0) {
      throw new TypeError('slideCount must be a non-negative safe integer');
    }
    if (
      options.generation !== undefined &&
      (!Number.isSafeInteger(options.generation) || options.generation <= 0)
    ) {
      throw new TypeError('generation must be a positive safe integer');
    }
  }

  /** Pull, optionally decode, and ACK exactly one complete slide. */
  async load(
    slideIndex: number,
    decode = true,
    timeoutMs?: number,
  ): Promise<Slide | undefined> {
    this.assertSlideIndex(slideIndex);
    const sessionId = this.nextSessionId++;
    const identity = {
      sessionId,
      operationId: sessionId,
      generation: this.options.generation ?? 1,
    };
    const session = new BoundedPullSession(this.options.transport, {
      ...identity,
      maxByteCredit: HARD_MAX_PPTX_SLIDE_JSON_BYTES,
      timeoutMs,
    });
    this.active.add(session);
    try {
      await this.options.open(slideIndex, identity, timeoutMs);
      const chunk = await pullWithCreditRetry(session);
      try {
        // JSON.parse is consumer acceptance for rendering pulls. Preflight pulls
        // deliberately skip it because the worker-side builder already parsed and
        // prepared the same transferred unit before this ACK.
        const slide = decode
          ? JSON.parse(new TextDecoder().decode(new Uint8Array(chunk.payload))) as Slide
          : undefined;
        await chunk.ack();
        return slide;
      } finally {
        // ACK releases producer staging; this separately drops the main-realm
        // transfer reference immediately instead of relying on a later GC.
        chunk.disposeTransferred();
      }
    } catch (error) {
      await session.cancel('request-error').catch(() => undefined);
      throw error;
    } finally {
      this.active.delete(session);
    }
  }

  /** Best-effort convergence used before terminating the owning worker. */
  cancelAll(): void {
    for (const session of this.active) {
      void session.cancel('closed').catch(() => undefined);
    }
    this.active.clear();
  }

  private assertSlideIndex(slideIndex: number): void {
    if (
      !Number.isSafeInteger(slideIndex) ||
      slideIndex < 0 ||
      slideIndex >= this.options.slideCount
    ) {
      throw new RangeError(
        `Slide index ${slideIndex} out of range (count: ${this.options.slideCount})`,
      );
    }
  }
}

export function isPptxSlidePullResponse(
  value: unknown,
): value is PullSessionResponse<ArrayBuffer, number> {
  return !!value && typeof value === 'object' &&
    (value as { protocol?: unknown }).protocol === PULL_SESSION_PROTOCOL;
}

async function pullWithCreditRetry(
  session: BoundedPullSession<ArrayBuffer, number>,
) {
  try {
    return await session.pull(PPTX_INITIAL_SLIDE_PULL_BYTES);
  } catch (error) {
    const required = requiredCredit(error);
    if (required === undefined) throw error;
    return session.pull(required);
  }
}

const REQUIRED_CREDIT = /^slide unit requires ([0-9]+) bytes but credit is ([0-9]+)$/u;

function requiredCredit(error: unknown): number | undefined {
  if (
    !error || typeof error !== 'object' ||
    (error as { code?: unknown }).code !== PULL_SESSION_INSUFFICIENT_CREDIT_CODE
  ) {
    return undefined;
  }
  const match = REQUIRED_CREDIT.exec(error instanceof Error ? error.message : String(error));
  if (!match) return undefined;
  const required = Number(match[1]);
  const offered = Number(match[2]);
  if (
    !Number.isSafeInteger(required) || required <= offered ||
    required > HARD_MAX_PPTX_SLIDE_JSON_BYTES ||
    offered !== PPTX_INITIAL_SLIDE_PULL_BYTES ||
    String(required) !== match[1] || String(offered) !== match[2]
  ) {
    return undefined;
  }
  return required;
}
