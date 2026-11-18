import { documentLimits } from './workbook';
import { requireDocument } from './workbook-validation';
import type { ReferenceBox } from './formula-references';

export type Axis = 'row' | 'column';
export interface AxisEdit {
  readonly axis: Axis;
  readonly start: number;
  readonly count: number;
  readonly delete: boolean;
  readonly lane?: readonly [first: number, last: number];
}

export function axisLimit(axis: Axis): number {
  return axis === 'row' ? documentLimits.rows : documentLimits.columns;
}

export function validateAxisEdit(edit: AxisEdit): void {
  const limit = axisLimit(edit.axis);
  requireDocument(Number.isInteger(edit.start) && Number.isInteger(edit.count) &&
    edit.start >= 1 && edit.count >= 1 && edit.start + edit.count - 1 <= limit, 'Invalid axis selection.');
  if (edit.lane) {
    const [first, last] = edit.lane;
    requireDocument(Number.isInteger(first) && Number.isInteger(last) && first >= 1 && last >= first &&
      last <= axisLimit(edit.axis === 'row' ? 'column' : 'row'), 'Invalid cell lane.');
  }
}

export function shiftIndex(edit: AxisEdit, index: number): number | undefined {
  if (index < edit.start) return index;
  if (edit.delete) return index < edit.start + edit.count ? undefined : index - edit.count;
  const next = index + edit.count;
  return next <= axisLimit(edit.axis) ? next : undefined;
}

export function shiftInterval(edit: AxisEdit, first: number, last: number): readonly [number, number] | undefined {
  if (last < edit.start) return [first, last];
  if (!edit.delete) {
    const end = shiftIndex(edit, last);
    return end === undefined ? undefined : [shiftIndex(edit, first)!, end];
  }
  const after = edit.start + edit.count;
  if (first >= edit.start && last < after) return undefined;
  return [first < edit.start ? first : Math.max(first, after) - edit.count,
    last < after ? edit.start - 1 : last - edit.count];
}

/** Undefined denotes a reference whose affected interval overflowed on insertion. */
export function shiftBox(edit: AxisEdit, box: ReferenceBox): readonly ReferenceBox[] | undefined {
  const vertical = edit.axis === 'row';
  const [first, last, laneFirst, laneLast] = vertical
    ? [box.top, box.bottom, box.left, box.right]
    : [box.left, box.right, box.top, box.bottom];
  const overlapFirst = Math.max(laneFirst, edit.lane?.[0] ?? laneFirst);
  const overlapLast = Math.min(laneLast, edit.lane?.[1] ?? laneLast);
  if (overlapFirst > overlapLast || last < edit.start) return [box];
  const span = shiftInterval(edit, first, last);
  if (!span && !edit.delete) return undefined;
  const parts: ReferenceBox[] = [];
  function add(a: number, b: number, c: number, d: number) {
    parts.push(vertical ? { top: a, bottom: b, left: c, right: d } : { top: c, bottom: d, left: a, right: b });
  }
  if (laneFirst < overlapFirst) add(first, last, laneFirst, overlapFirst - 1);
  if (overlapLast < laneLast) add(first, last, overlapLast + 1, laneLast);
  if (span) add(...span, overlapFirst, overlapLast);
  return parts;
}
