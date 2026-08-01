import { describe, expect, it } from 'vitest';
import type { WorkerBridgeTransport } from '@silurus/ooxml-core';
import {
  PULL_SESSION_INSUFFICIENT_CREDIT_CODE,
  PULL_SESSION_PROTOCOL,
  type PullSessionCommand,
  type PullSessionResponse,
} from '@silurus/ooxml-core/worker';
import {
  PPTX_INITIAL_SLIDE_PULL_BYTES,
  PptxSlidePullClient,
} from './slide-pull-client.js';
import type { Slide } from './types.js';

type Response = PullSessionResponse<ArrayBuffer, number>;

function slide(index: number): Slide {
  return {
    index,
    slideNumber: index + 1,
    background: null,
    elements: [],
  };
}

class FakeTransport implements WorkerBridgeTransport<Response> {
  readonly commands: PullSessionCommand<number>[] = [];
  nextRequestId = 1;
  handler: (command: PullSessionCommand<number>) => Response = () => {
    throw new Error('no handler');
  };

  request(build: (id: number) => unknown): Promise<Response> {
    const command = build(this.nextRequestId++) as PullSessionCommand<number>;
    this.commands.push(command);
    return Promise.resolve(this.handler(command));
  }

  forgetOrphaned(): void {}
  terminate(): void {}
}

function response(
  command: PullSessionCommand<number>,
  value: Record<string, unknown>,
): Response {
  return {
    protocol: PULL_SESSION_PROTOCOL,
    sessionId: command.sessionId,
    operationId: command.operationId,
    generation: command.generation,
    requestId: command.requestId,
    ...value,
  } as Response;
}

describe('PptxSlidePullClient', () => {
  it('retries coded insufficient credit with the same identity and sequence', async () => {
    const transport = new FakeTransport();
    const required = PPTX_INITIAL_SLIDE_PULL_BYTES + 17;
    transport.handler = (command) => {
      if (command.kind === 'pull' && command.byteCredit === PPTX_INITIAL_SLIDE_PULL_BYTES) {
        return response(command, {
          kind: 'error',
          error: {
            message: `slide unit requires ${required} bytes but credit is ${PPTX_INITIAL_SLIDE_PULL_BYTES}`,
            errorName: 'RangeError',
            code: PULL_SESSION_INSUFFICIENT_CREDIT_CODE,
          },
        });
      }
      if (command.kind === 'pull') {
        const payload = new TextEncoder().encode(JSON.stringify(slide(2))).buffer;
        return response(command, {
          kind: 'chunk',
          sequence: command.sequence,
          byteLength: payload.byteLength,
          done: true,
          payload,
        });
      }
      return response(command, { kind: 'accepted', command: command.kind });
    };
    const opens: unknown[] = [];
    const client = new PptxSlidePullClient({
      slideCount: 3,
      transport,
      open: async (index, identity) => { opens.push({ index, identity }); },
    });

    await expect(client.load(2)).resolves.toEqual(slide(2));
    const pulls = transport.commands.filter(
      (command): command is Extract<PullSessionCommand<number>, { kind: 'pull' }> =>
        command.kind === 'pull',
    );
    expect(pulls).toHaveLength(2);
    expect(pulls.map(({ sequence, sessionId, operationId, generation, byteCredit }) => ({
      sequence,
      sessionId,
      operationId,
      generation,
      byteCredit,
    }))).toEqual([
      { sequence: 0, sessionId: 1, operationId: 1, generation: 1, byteCredit: PPTX_INITIAL_SLIDE_PULL_BYTES },
      { sequence: 0, sessionId: 1, operationId: 1, generation: 1, byteCredit: required },
    ]);
    expect(opens).toEqual([{
      index: 2,
      identity: { sessionId: 1, operationId: 1, generation: 1 },
    }]);
  });

  it('can ACK preflight bytes without decoding a complete Slide in Window', async () => {
    const transport = new FakeTransport();
    transport.handler = (command) => command.kind === 'pull'
      ? response(command, {
          kind: 'chunk',
          sequence: command.sequence,
          byteLength: 8,
          done: true,
          payload: new TextEncoder().encode('not-json').buffer,
        })
      : response(command, { kind: 'accepted', command: command.kind });
    const client = new PptxSlidePullClient({
      slideCount: 1,
      transport,
      open: async () => undefined,
    });

    await expect(client.load(0, false)).resolves.toBeUndefined();
    expect(transport.commands.map((command) => command.kind)).toEqual(['pull', 'ack']);
  });

  it('scopes a timeout to the requested load instead of retaining it for later slides', async () => {
    const transport = new FakeTransport();
    transport.handler = (command) => command.kind === 'pull'
      ? response(command, {
          kind: 'chunk',
          sequence: command.sequence,
          byteLength: 8,
          done: true,
          payload: new TextEncoder().encode('not-json').buffer,
        })
      : response(command, { kind: 'accepted', command: command.kind });
    const timeouts: Array<number | undefined> = [];
    const client = new PptxSlidePullClient({
      slideCount: 2,
      transport,
      open: async (_index, _identity, timeoutMs) => { timeouts.push(timeoutMs); },
    });

    await client.load(0, false, 1234);
    await client.load(1, false);
    expect(timeouts).toEqual([1234, undefined]);
  });

  it('cancels instead of ACKing when consumer JSON acceptance fails', async () => {
    const transport = new FakeTransport();
    transport.handler = (command) => command.kind === 'pull'
      ? response(command, {
          kind: 'chunk',
          sequence: command.sequence,
          byteLength: 1,
          done: true,
          payload: new Uint8Array([0xff]).buffer,
        })
      : response(command, { kind: 'accepted', command: command.kind });
    const client = new PptxSlidePullClient({
      slideCount: 1,
      transport,
      open: async () => undefined,
    });

    await expect(client.load(0)).rejects.toThrow();
    expect(transport.commands.some((command) => command.kind === 'ack')).toBe(false);
    expect(transport.commands.some((command) => command.kind === 'cancel')).toBe(true);
  });
});
