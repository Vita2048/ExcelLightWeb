import type { AxisDimension, Coordinate } from '../core/workbook';
import { columnLetters } from '../core/formula-references';

export const gridMetrics = Object.freeze({ row: 26, column: 96, rowHeader: 44, columnHeader: 26 });
/**
 * Largest content extent mapped 1:1 onto a native scroll offset.
 * Chromium stops smooth scrolling once a box reaches 2^24 px, so the spacer stays one pixel under this.
 */
export const scrollCap = 16_777_216;
export interface Track {
  readonly index: number;
  readonly start: number;
  readonly size: number;
}

interface Segment {
  readonly index: number;
  readonly count: number;
  readonly start: number;
  readonly size: number;
  readonly ordinal: number;
}

export interface HiddenSpan {
  readonly start: number;
  readonly end: number;
}

function mergeSpans(spans: readonly HiddenSpan[], lastIndex: number): HiddenSpan[] {
  const clipped = spans
    .map(span => ({ start: Math.max(1, Math.ceil(span.start)), end: Math.min(lastIndex, Math.floor(span.end)) }))
    .filter(span => span.start <= span.end)
    .sort((left, right) => left.start - right.start || left.end - right.end);
  const merged: HiddenSpan[] = [];
  for (const span of clipped) {
    const previous = merged.at(-1);
    if (previous && span.start <= previous.end + 1) {
      merged[merged.length - 1] = { start: previous.start, end: Math.max(previous.end, span.end) };
    } else merged.push({ start: span.start, end: span.end });
  }
  return merged;
}

/** Removes inclusive spans from visible runs and reflows pixel starts. A span hides a custom-size row. */
function withoutSpans(segments: readonly Segment[], spans: readonly HiddenSpan[]): Segment[] {
  const next: Segment[] = [];
  let origin = 0;
  let ordinal = 0;
  const push = (index: number, run: number, length: number) => {
    if (run <= 0) return;
    const previous = next.at(-1);
    if (previous && previous.index + previous.count === index && previous.size === length) {
      next[next.length - 1] = { ...previous, count: previous.count + run };
      origin += run * length;
      ordinal += run;
      return;
    }
    next.push({ index, count: run, start: origin, size: length, ordinal });
    origin += run * length;
    ordinal += run;
  };
  let spanIndex = 0;
  for (const segment of segments) {
    let cursor = segment.index;
    const end = segment.index + segment.count - 1;
    while (spanIndex < spans.length && spans[spanIndex]!.end < cursor) spanIndex++;
    let look = spanIndex;
    while (cursor <= end && look < spans.length && spans[look]!.start <= end) {
      const span = spans[look]!;
      if (span.start > cursor) push(cursor, Math.min(end, span.start - 1) - cursor + 1, segment.size);
      cursor = Math.max(cursor, span.end + 1);
      if (span.end < end) look++;
      else break;
    }
    if (cursor <= end) push(cursor, end - cursor + 1, segment.size);
  }
  return next;
}

export function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

/**
 * Pixel length of the scroll spacer, including the sticky header.
 * The length stays strictly under `cap` (2^24). At that size Chromium clamps the box and
 * ignores slow scrolling, which hides the last rows even when scrollTop is at its maximum.
 */
export function scrollSpacer(content: number, header = 0, cap = scrollCap): number {
  const span = Math.max(0, header) + Math.max(0, content);
  return Math.min(span, Math.max(0, cap - 1));
}

/**
 * Scroll-mapping cap implied by the browser's used scroll range.
 * `used` is `scrollHeight` or `scrollWidth`. When the browser shortens the spacer, the map follows
 * that shorter range so the last row still meets the bottom of the window.
 */
export function capFromScrollport(used: number, header: number, content: number): number {
  if (used > header) return used - header;
  return Math.min(Math.max(0, content), Math.max(0, scrollCap - header));
}

/**
 * Width of the row-number column. `textWidth` is the measured width of the last row's label
 * at the header font. The column stays at least the historical 44px header.
 */
export function rowHeaderWidth(textWidth: number, scale = 1): number {
  return Math.max(gridMetrics.rowHeader * scale, Math.ceil(textWidth + 12 * scale));
}

/**
 * Native scroll offset to content pixels. Identity while `content` fits under the cap.
 * `port` is the visible body (scrollport minus the sticky header). When the sheet is compressed,
 * scroll 0 is the start and the maximum native scroll places the end of the sheet on the bottom of `port`.
 */
export function contentOffset(scrollPosition: number, content: number, cap = scrollCap, port = 0): number {
  if (!(content > cap) || cap === 0) return scrollPosition;
  const maxOffset = Math.max(0, content - port);
  const maxScroll = Math.max(0, cap - port);
  if (maxScroll === 0 || maxOffset === 0) return 0;
  return clamp(scrollPosition, 0, maxScroll) * maxOffset / maxScroll;
}

/** Content pixels to a native scroll offset. Identity while `content` fits under the cap. */
export function scrollTopFor(offset: number, content: number, cap = scrollCap, port = 0): number {
  if (!(content > cap) || content === 0) return offset;
  const maxOffset = Math.max(0, content - port);
  const maxScroll = Math.max(0, cap - port);
  if (maxOffset === 0 || maxScroll === 0) return 0;
  return clamp(offset, 0, maxOffset) * maxScroll / maxOffset;
}

/**
 * Native scroll offset for a content position, starting from the current native offset.
 * A compressed move that would stay on the same device pixel steps one pixel so the browser scrolls.
 */
export function placeScroll(offset: number, content: number, current: number, cap = scrollCap, port = 0): number {
  const next = scrollTopFor(offset, content, cap, port);
  if (!(content > cap) || content === 0) return next;
  const from = contentOffset(current, content, cap, port);
  if (offset > from) {
    const maxScroll = Math.max(0, cap - port);
    return Math.min(maxScroll, Math.max(current + 1, Math.ceil(next)));
  }
  if (offset < from) return Math.max(0, Math.min(current - 1, Math.floor(next)));
  return current;
}

export class GridAxis {
  readonly total: number;
  readonly frozenSize: number;
  readonly frozen: number;
  /** Uniform visible runs. An empty sheet is one run, independent of the row count. */
  readonly runs: number;
  readonly #segments: readonly Segment[];
  readonly #visible: number;

  constructor(count: number, size: number, dimensions: readonly AxisDimension[], scale = 1, frozen = 0,
    hiddenSpans: readonly HiddenSpan[] = []) {
    const lastIndex = Math.floor(count);
    const byIndex = new Map<number, AxisDimension>();
    for (const dimension of dimensions) {
      if (dimension.index >= 1 && dimension.index <= lastIndex) byIndex.set(dimension.index, dimension);
    }
    const ordered = [...byIndex.values()].sort((a, b) => a.index - b.index);
    const segments: Segment[] = [];
    let cursor = 1;
    let origin = 0;
    let ordinal = 0;
    const push = (index: number, run: number, length: number) => {
      if (run <= 0) return;
      const previous = segments.at(-1);
      if (previous && previous.index + previous.count === index && previous.size === length) {
        segments[segments.length - 1] = { ...previous, count: previous.count + run };
        origin += run * length;
        ordinal += run;
        return;
      }
      segments.push({ index, count: run, start: origin, size: length, ordinal });
      origin += run * length;
      ordinal += run;
    };
    for (const dimension of ordered) {
      if (dimension.index > cursor) push(cursor, dimension.index - cursor, size * scale);
      if (!dimension.hidden) push(dimension.index, 1, (dimension.size ?? size) * scale);
      cursor = dimension.index + 1;
    }
    if (cursor <= lastIndex) push(cursor, lastIndex - cursor + 1, size * scale);
    const hidden = mergeSpans(hiddenSpans, lastIndex);
    const visibleSegments = hidden.length ? withoutSpans(segments, hidden) : segments;
    const tail = visibleSegments.at(-1);
    this.#segments = visibleSegments;
    this.#visible = tail ? tail.ordinal + tail.count : 0;
    this.runs = visibleSegments.length;
    this.total = tail ? tail.start + tail.count * tail.size : 0;
    this.frozen = Math.min(count, frozen);
    const boundary = this.trackAtOrAfter(Math.floor(this.frozen) + 1);
    this.frozenSize = boundary ? boundary.start : this.total;
  }

  get first(): Track | undefined {
    return this.#trackFromOrdinal(0);
  }

  get last(): Track | undefined {
    return this.#trackFromOrdinal(this.#visible - 1);
  }

  #segmentAt(position: number): Segment | undefined {
    let low = 0;
    let high = this.#segments.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      const segment = this.#segments[middle]!;
      if (segment.start + segment.count * segment.size <= position) low = middle + 1;
      else high = middle;
    }
    return this.#segments[low];
  }

  #segmentContaining(index: number): Segment | undefined {
    let low = 0;
    let high = this.#segments.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      const segment = this.#segments[middle]!;
      if (segment.index + segment.count <= index) low = middle + 1;
      else high = middle;
    }
    const segment = this.#segments[low];
    return segment && index >= segment.index ? segment : undefined;
  }

  #segmentAfter(index: number): Segment | undefined {
    let low = 0;
    let high = this.#segments.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      if (this.#segments[middle]!.index <= index) low = middle + 1;
      else high = middle;
    }
    return this.#segments[low];
  }

  #trackIn(segment: Segment, offset: number): Track {
    return {
      index: segment.index + offset,
      start: segment.start + offset * segment.size,
      size: segment.size,
    };
  }

  #trackFromOrdinal(ordinal: number): Track | undefined {
    if (ordinal < 0 || ordinal >= this.#visible) return undefined;
    let low = 0;
    let high = this.#segments.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      const segment = this.#segments[middle]!;
      if (segment.ordinal + segment.count <= ordinal) low = middle + 1;
      else high = middle;
    }
    const segment = this.#segments[low]!;
    return this.#trackIn(segment, ordinal - segment.ordinal);
  }

  #ordinalAt(position: number): number {
    const segment = this.#segmentAt(position);
    if (!segment || segment.size === 0) return this.#visible;
    if (position < segment.start) return segment.ordinal;
    const offset = Math.floor((position - segment.start) / segment.size);
    return offset >= segment.count ? this.#visible : segment.ordinal + offset;
  }

  at(position: number): Track | undefined {
    if (position < 0) return undefined;
    const ordinal = this.#ordinalAt(position);
    return ordinal < this.#visible ? this.#trackFromOrdinal(ordinal) : undefined;
  }

  visible(start: number, length: number): readonly Track[] {
    if (length <= 0 || this.#visible === 0) return [];
    const from = this.#ordinalAt(start);
    const end = Math.min(this.#visible, this.#ordinalAt(start + length - 0.000001) + 1);
    if (from >= end) return [];
    const tracks: Track[] = [];
    for (const segment of this.#segments) {
      const segmentEnd = segment.ordinal + segment.count;
      if (segmentEnd <= from) continue;
      if (segment.ordinal >= end) break;
      const first = Math.max(from, segment.ordinal);
      const stop = Math.min(end, segmentEnd);
      for (let ordinal = first; ordinal < stop; ordinal++) tracks.push(this.#trackIn(segment, ordinal - segment.ordinal));
    }
    return tracks;
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
    const segment = this.#segmentContaining(index);
    return segment ? this.#trackIn(segment, index - segment.index) : undefined;
  }

  trackAtOrAfter(index: number): Track | undefined {
    const segment = this.#segmentContaining(index) ?? this.#segmentAfter(index);
    return segment ? this.#trackIn(segment, Math.max(0, index - segment.index)) : undefined;
  }

  step(index: number, count: number): number {
    if (this.#visible === 0) return index;
    const segment = this.#segmentContaining(index);
    const ordinal = segment
      ? segment.ordinal + (index - segment.index)
      : (this.#segmentAfter(index)?.ordinal ?? this.#visible) - (count > 0 ? 1 : 0);
    return this.#trackFromOrdinal(clamp(ordinal + count, 0, this.#visible - 1))!.index;
  }

  reveal(index: number, scroll: number, extent: number): number {
    const track = this.track(index);
    if (!track || track.index <= this.frozen || extent <= this.frozenSize) return scroll;
    const end = track.start + Math.min(track.size, extent - this.frozenSize);
    return clamp(scroll, Math.max(0, end - extent), track.start - this.frozenSize);
  }

  span(first: number, last: number): number {
    if (first === last) return this.track(first)?.size ?? 0;
    const start = this.track(first);
    const end = this.track(last);
    if (start && end) return end.start + end.size - start.start;
    let origin = 0;
    let finish = 0;
    let seen = false;
    for (const segment of this.#segments) {
      const segmentLast = segment.index + segment.count - 1;
      if (segmentLast < first || segment.index > last) continue;
      const from = Math.max(segment.index, first);
      const to = Math.min(segmentLast, last);
      const fromStart = segment.start + (from - segment.index) * segment.size;
      const toEnd = segment.start + (to - segment.index + 1) * segment.size;
      if (!seen) origin = fromStart;
      finish = toEnd;
      seen = true;
    }
    return seen ? finish - origin : 0;
  }
}

export function columnName(column: number): string {
  return columnLetters(column);
}

export function address(point: Coordinate): string {
  return columnName(point.column) + point.row;
}
