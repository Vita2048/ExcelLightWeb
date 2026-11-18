import type { AxisDimension, Coordinate } from '../core/workbook';
import { columnLetters } from '../core/formula-references';

export const gridMetrics = Object.freeze({ row: 26, column: 96, rowHeader: 44, columnHeader: 26 });
export interface Track {
  readonly index: number;
  readonly start: number;
  readonly size: number;
}

export function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

export class GridAxis {
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

  screen(track: Track, scroll: number): number {
    return track.start - (track.index <= this.frozen ? 0 : scroll);
  }

  atScreen(position: number, scroll: number): Track | undefined {
    return this.at(position + (position >= this.frozenSize ? scroll : 0));
  }

  projected(scroll: number, extent: number): Track[] {
    const pinned = this.visible(0, Math.min(extent, this.frozenSize)).filter(track => track.index <= this.frozen);
    const visible = this.visible(this.frozenSize + scroll, extent - this.frozenSize).filter(track => track.index > this.frozen);
    return [...pinned, ...visible.map(track => ({ ...track, start: this.screen(track, scroll) }))];
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
    return this.tracks[clamp(offset + count, 0, this.tracks.length - 1)]?.index ?? index;
  }

  reveal(index: number, scroll: number, extent: number): number {
    const track = this.track(index);
    if (!track || track.index <= this.frozen || extent <= this.frozenSize) return scroll;
    const end = track.start + Math.min(track.size, extent - this.frozenSize);
    return clamp(scroll, Math.max(0, end - extent), track.start - this.frozenSize);
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

export function columnName(column: number): string {
  return columnLetters(column);
}

export function address(point: Coordinate): string {
  return columnName(point.column) + point.row;
}
