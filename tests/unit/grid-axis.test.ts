import { describe, expect, it } from 'vitest';
import type { AxisDimension } from '../../src/core/workbook';
import {
  capFromScrollport, contentOffset, GridAxis, placeScroll, rowHeaderWidth, scrollCap, scrollSpacer, scrollTopFor, type Track,
} from '../../src/grid/grid-geometry';

/** The previous per-track axis, kept as an oracle for small sheets. */
class DenseAxis {
  readonly tracks: readonly Track[];
  readonly total: number;
  readonly frozenSize: number;
  readonly frozen: number;
  readonly #byIndex: ReadonlyMap<number, number>;

  constructor(count: number, size: number, dimensions: readonly AxisDimension[], scale = 1, frozen = 0) {
    const dimensionsByIndex = new Map(dimensions.map(dimension => [dimension.index, dimension]));
    const tracks: Track[] = [];
    const byIndex = new Map<number, number>();
    let offset = 0;
    for (let index = 1; index <= count; index++) {
      const dimension = dimensionsByIndex.get(index);
      if (dimension?.hidden) continue;
      const length = (dimension?.size ?? size) * scale;
      byIndex.set(index, tracks.length);
      tracks.push({ index, start: offset, size: length });
      offset += length;
    }
    this.tracks = tracks;
    this.total = offset;
    this.frozen = Math.min(count, frozen);
    const lastFrozen = tracks.findIndex(track => track.index > this.frozen);
    this.frozenSize = lastFrozen < 0 ? offset : tracks[lastFrozen]?.start ?? 0;
    this.#byIndex = byIndex;
  }

  #offsetAt(position: number): number {
    let low = 0;
    let high = this.tracks.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      const track = this.tracks[middle]!;
      if (track.start + track.size <= position) low = middle + 1;
      else high = middle;
    }
    return low;
  }

  at(position: number): Track | undefined {
    return position >= 0 ? this.tracks[this.#offsetAt(position)] : undefined;
  }

  visible(start: number, length: number): readonly Track[] {
    if (length <= 0) return [];
    return this.tracks.slice(this.#offsetAt(start), this.#offsetAt(start + length - 0.000001) + 1);
  }

  track(index: number): Track | undefined {
    return this.tracks[this.#byIndex.get(index) ?? -1];
  }

  step(index: number, count: number): number {
    let offset = this.#byIndex.get(index);
    if (offset === undefined) {
      const following = this.tracks.findIndex(track => track.index > index);
      offset = following < 0 ? this.tracks.length : following;
      if (count > 0) offset--;
    }
    const clamped = Math.min(Math.max(offset + count, 0), this.tracks.length - 1);
    return this.tracks[clamped]?.index ?? index;
  }

  reveal(index: number, scroll: number, extent: number): number {
    const track = this.track(index);
    if (!track || track.index <= this.frozen || extent <= this.frozenSize) return scroll;
    const end = track.start + Math.min(track.size, extent - this.frozenSize);
    return Math.min(Math.max(scroll, Math.max(0, end - extent)), track.start - this.frozenSize);
  }

  span(first: number, last: number): number {
    if (first === last) return this.track(first)?.size ?? 0;
    const a = this.track(first), b = this.track(last);
    if (a && b) return b.start + b.size - a.start;
    const visible = this.tracks.filter(track => track.index >= first && track.index <= last);
    const end = visible.at(-1);
    return end ? end.start + end.size - visible[0]!.start : 0;
  }
}

function expectSame(count: number, size: number, dimensions: readonly AxisDimension[], scale = 1, frozen = 0) {
  const dense = new DenseAxis(count, size, dimensions, scale, frozen);
  const sparse = new GridAxis(count, size, dimensions, scale, frozen);
  expect(sparse.total).toBe(dense.total);
  expect(sparse.frozen).toBe(dense.frozen);
  expect(sparse.frozenSize).toBe(dense.frozenSize);
  expect(sparse.first).toEqual(dense.tracks[0]);
  expect(sparse.last).toEqual(dense.tracks.at(-1));
  for (const position of [0, 0.1, 25.9, 26, 39, 77.9, 78, 100, dense.total - 0.001, dense.total, -1]) {
    expect(sparse.at(position), `at ${position}`).toEqual(dense.at(position));
  }
  expect(sparse.visible(0, 80).map(track => track.index)).toEqual(dense.visible(0, 80).map(track => track.index));
  expect(sparse.visible(40, 40)).toEqual(dense.visible(40, 40));
  for (let index = 0; index <= count + 1; index++) {
    expect(sparse.track(index), `track ${index}`).toEqual(dense.track(index));
    expect(sparse.step(index, 1), `step+ ${index}`).toBe(dense.step(index, 1));
    expect(sparse.step(index, -1), `step- ${index}`).toBe(dense.step(index, -1));
    expect(sparse.step(index, 8), `step page ${index}`).toBe(dense.step(index, 8));
    expect(sparse.reveal(index, 0, 90), `reveal ${index}`).toBe(dense.reveal(index, 0, 90));
  }
  expect(sparse.span(2, 4)).toBe(dense.span(2, 4));
  expect(sparse.span(3, 3)).toBe(dense.span(3, 3));
  expect(sparse.span(1, count)).toBe(dense.span(1, count));
  const projected = sparse.projected(30, 120);
  const densePinned = dense.visible(0, Math.min(120, dense.frozenSize)).filter(track => track.index <= dense.frozen);
  const denseBody = dense.visible(dense.frozenSize + 30, 120 - dense.frozenSize).filter(track => track.index > dense.frozen)
    .map(track => ({ ...track, start: track.start - (track.index <= dense.frozen ? 0 : 30) }));
  expect(projected).toEqual([...densePinned, ...denseBody]);
}

describe('sparse grid axis', () => {
  const mixed: AxisDimension[] = [
    { index: 4, size: 38 },
    { index: 3, hidden: true },
    { index: 4, size: 40 },
    { index: 8, hidden: true },
  ];

  it('matches the dense axis for hidden rows, custom sizes, zoom and frozen panes', () => {
    expectSame(12, 26, mixed, 1.5, 1);
    expectSame(12, 26, mixed, 1.25, 4);
    expectSame(6, 26, [{ index: 3, hidden: true }, { index: 4, size: 38 }], 1.5, 0);
  });

  it('addresses the last row of an Excel-sized axis without allocating one run per row', () => {
    const started = performance.now();
    const axis = new GridAxis(1_048_576, 26, [{ index: 10, hidden: true }, { index: 500_000, size: 40 }]);
    expect(performance.now() - started).toBeLessThan(50);
    expect(axis.runs).toBe(4);
    expect(axis.total).toBe(9 * 26 + (500_000 - 11) * 26 + 40 + (1_048_576 - 500_000) * 26);
    expect(axis.at(0)?.index).toBe(1);
    expect(axis.track(1_048_576)?.index).toBe(1_048_576);
    expect(axis.last?.index).toBe(1_048_576);
    expect(axis.at(axis.total - 1)?.index).toBe(1_048_576);
    expect(axis.at(axis.total)).toBeUndefined();
    expect(axis.reveal(1_048_576, 0, 200)).toBe(axis.track(1_048_576)!.start + 26 - 200);
    expect(axis.visible(axis.total - 100, 100).map(track => track.index)).toEqual([1_048_573, 1_048_574, 1_048_575, 1_048_576]);
    expect(axis.step(1, 1)).toBe(2);
    expect(axis.step(9, 1)).toBe(11);
    expect(axis.step(10, 1)).toBe(11);
    expect(axis.step(10, -1)).toBe(9);
  });

  it('punches a hidden span out of an Excel-sized axis without one run per row', () => {
    const started = performance.now();
    const axis = new GridAxis(1_048_576, 26, [{ index: 4, size: 40 }], 1, 1, [{ start: 2, end: 1_048_575 }]);
    expect(performance.now() - started).toBeLessThan(50);
    expect(axis.runs).toBe(2);
    expect(axis.track(1)?.index).toBe(1);
    expect(axis.track(4)).toBeUndefined();
    expect(axis.track(1_048_576)?.index).toBe(1_048_576);
    expect(axis.step(1, 1)).toBe(1_048_576);
    const hole = new GridAxis(100, 26, [{ index: 10, size: 40 }], 1, 0, [{ start: 1, end: 9 }, { start: 11, end: 100 }]);
    expect(hole.track(10)?.size).toBe(40);
    expect(hole.runs).toBe(1);
    expect(hole.total).toBe(40);
  });

  it('keeps an empty sheet to a single run', () => {
    const axis = new GridAxis(1_048_576, 26, []);
    expect(axis.runs).toBe(1);
    expect(axis.total).toBe(1_048_576 * 26);
    expect(axis.track(1)?.start).toBe(0);
    expect(axis.track(1_048_576)?.start).toBe((1_048_576 - 1) * 26);
  });
});

describe('scroll mapping', () => {
  it('leaves scroll offsets unchanged while the sheet fits under the cap', () => {
    expect(scrollSpacer(1_000 * 26)).toBe(26_000);
    expect(contentOffset(400, 26_000)).toBe(400);
    expect(scrollTopFor(400, 26_000)).toBe(400);
  });

  it('compresses a sheet taller than the cap and round-trips the fraction', () => {
    const content = 32_000_000;
    expect(scrollSpacer(content)).toBe(scrollCap - 1);
    expect(contentOffset(scrollCap / 2, content)).toBe(content / 2);
    expect(scrollTopFor(content / 2, content)).toBe(scrollCap / 2);
    expect(contentOffset(scrollTopFor(123_456, content), content)).toBeCloseTo(123_456, 5);
  });

  it('steps a compressed move that would round to the same device pixel and keeps revealed edges unclipped', () => {
    const content = 1_048_576 * 26;
    const port = 495;
    const cap = scrollCap - 1 - 26;
    expect(scrollTopFor(1, content)).toBeLessThan(1);
    expect(placeScroll(1, content, 0)).toBe(1);
    expect(placeScroll(0, content, 10)).toBe(0);
    expect(placeScroll(400, 26_000, 0)).toBe(400);
    const forward = placeScroll(25, content, 0, cap, port);
    expect(Number.isInteger(forward)).toBe(true);
    expect(contentOffset(forward, content, cap, port)).toBeGreaterThanOrEqual(25);
    const unchanged = placeScroll(contentOffset(forward, content, cap, port), content, forward, cap, port);
    expect(unchanged).toBe(forward);
    const backward = placeScroll(26, content, 32, cap, port);
    expect(Number.isInteger(backward)).toBe(true);
    expect(contentOffset(backward, content, cap, port)).toBeLessThanOrEqual(26);
  });

  it('places the last row inside the scrollport when the sheet is taller than the cap', () => {
    const content = 1_048_576 * 26;
    const port = 480;
    const header = 26;
    const used = scrollSpacer(content, header);
    expect(used).toBe(scrollCap - 1);
    const cap = capFromScrollport(used, header, content);
    const maxScroll = used - (port + header);
    expect(contentOffset(maxScroll, content, cap, port) + port).toBeCloseTo(content, 4);
    expect(scrollTopFor(content - port, content, cap, port)).toBeCloseTo(maxScroll, 4);
    const axis = new GridAxis(1_048_576, 26, []);
    const shown = axis.reveal(1_048_576, 0, port);
    expect(shown + port).toBeCloseTo(content, 4);
    expect(scrollTopFor(shown, content, cap, port)).toBeLessThanOrEqual(maxScroll + 0.001);
    const clamped = scrollCap - 64;
    const clampedCap = capFromScrollport(clamped, header, content);
    const clampedScroll = clamped - (port + header);
    expect(contentOffset(clampedScroll, content, clampedCap, port) + port).toBeCloseTo(content, 4);
  });

  it('sizes the row header from the last row number', () => {
    expect(rowHeaderWidth(20)).toBe(44);
    expect(rowHeaderWidth(56)).toBe(68);
    expect(rowHeaderWidth(112, 2)).toBe(136);
    expect(rowHeaderWidth(20, 2)).toBe(88);
  });
});
