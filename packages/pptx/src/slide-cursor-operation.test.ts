import { describe, expect, it, vi } from 'vitest';
import type { OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import {
  loadPptxSlideFromCursor,
  type PptxSlideCursorArchive,
} from './slide-cursor-operation.js';
import type { Slide } from './types.js';

const usage: OoxmlResourceUsageSnapshot = {
  archiveEntryCount: 3,
  declaredInflatedBytes: 30,
  distinctInflatedBytes: 20,
  operationInflatedBytes: 10,
};

function slide(index = 0): Slide {
  return {
    index,
    slideNumber: index + 1,
    background: null,
    elements: [],
  };
}

function archiveFor(value = slide()) {
  const events: string[] = [];
  const archive: PptxSlideCursorArchive = {
    pull_slide: vi.fn(() => {
      events.push('pull');
      return new TextEncoder().encode(JSON.stringify(value));
    }),
    slide_cursor_resource_usage: vi.fn(() => {
      events.push('usage');
      return new TextEncoder().encode(JSON.stringify(usage));
    }),
    acknowledge_slide: vi.fn(() => { events.push('ack'); }),
    cancel_slide: vi.fn(() => { events.push('cancel'); }),
    close_presentation_session: vi.fn(),
  };
  return { archive, events };
}

describe('loadPptxSlideFromCursor', () => {
  it('prepares compact consumer state before Rust ACK and publishes it after ACK', () => {
    const { archive, events } = archiveFor();
    const loaded = loadPptxSlideFromCursor(
      (operation) => operation(archive),
      0,
      { operationId: 7, generation: 3 },
      (_index, _slide, checkpoint) => {
        events.push('prepare');
        expect(checkpoint).toEqual(usage);
        return {
          rollback: () => events.push('rollback'),
          commit: () => events.push('commit'),
        };
      },
    );

    expect(loaded).toEqual(slide());
    expect(events).toEqual(['pull', 'usage', 'prepare', 'ack', 'commit']);
    expect(archive.acknowledge_slide).toHaveBeenCalledWith(7, 3);
    expect(archive.cancel_slide).not.toHaveBeenCalled();
  });

  it('rolls back and cancels when Rust ACK fails', () => {
    const { archive, events } = archiveFor();
    vi.mocked(archive.acknowledge_slide).mockImplementation(() => {
      events.push('ack');
      throw new Error('ack failed');
    });

    expect(() => loadPptxSlideFromCursor(
      (operation) => operation(archive),
      0,
      { operationId: 1, generation: 1 },
      () => ({
        rollback: () => events.push('rollback'),
        commit: () => events.push('commit'),
      }),
    )).toThrow('ack failed');
    expect(events).toEqual(['pull', 'usage', 'ack', 'rollback', 'cancel']);
  });

  it('rejects malformed usage before ACK and preserves cancellation', () => {
    const { archive } = archiveFor();
    vi.mocked(archive.slide_cursor_resource_usage).mockReturnValue(
      new TextEncoder().encode('{"archiveEntryCount":-1}'),
    );

    expect(() => loadPptxSlideFromCursor(
      (operation) => operation(archive),
      0,
      { operationId: 1, generation: 1 },
      () => undefined,
    )).toThrow(/usage checkpoint is invalid/);
    expect(archive.acknowledge_slide).not.toHaveBeenCalled();
    expect(archive.cancel_slide).toHaveBeenCalledOnce();
  });
});
