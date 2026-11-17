import type { AxisDimension, Coordinate } from '../core/workbook';

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
  readonly #byIndex: ReadonlyMap<number, number>;

  constructor(count: number, size: number, dimensions: readonly AxisDimension[], scale = 1) {
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
    return this.tracks[clamp(offset + count, 0, this.tracks.length - 1)]?.index ?? index;
  }

  reveal(index: number, scroll: number, extent: number): number {
    const track = this.track(index);
    if (!track || extent <= 0) return scroll;
    const end = track.start + Math.min(track.size, extent);
    return clamp(scroll, Math.max(0, end - extent), track.start);
  }

  span(first: number, last: number): number {
    const visible = this.tracks.filter(track => track.index >= first && track.index <= last);
    const end = visible.at(-1);
    return end ? end.start + end.size - visible[0]!.start : 0;
  }
}

export function columnName(column: number): string {
  let label = '';
  for (let remaining = column; remaining > 0; remaining = Math.floor((remaining - 1) / 26)) {
    label = String.fromCharCode(65 + (remaining - 1) % 26) + label;
  }
  return label;
}

export function address(point: Coordinate): string {
  return columnName(point.column) + point.row;
}
