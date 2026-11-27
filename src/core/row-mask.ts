/** Hidden rows inside a filter span, stored as the smaller of the two outcomes. */
export interface RowSpan {
  readonly start: number;
  readonly end: number;
}

export interface MaskCriterion {
  readonly column: number;
  readonly allowed: ReadonlySet<string>;
}

export interface MaskCell {
  readonly row: number;
  readonly column: number;
  readonly key: string;
}

/**
 * A filter span hides either a few listed rows (`hidden`) or every row except a few (`shown`).
 * Membership, counts and spans stay proportional to the exception list.
 */
export class RowMask {
  readonly mode: 'none' | 'hidden' | 'shown';
  /** First row the mask applies to, inclusive. */
  readonly from: number;
  /** Last row the mask applies to, inclusive. */
  readonly to: number;
  readonly #rows: readonly number[];

  private constructor(mode: 'none' | 'hidden' | 'shown', from: number, to: number, rows: readonly number[]) {
    this.mode = mode;
    this.from = from;
    this.to = to;
    this.#rows = rows;
  }

  private static readonly empty = new RowMask('none', 1, 0, []);

  static none(): RowMask {
    return RowMask.empty;
  }

  /** Rows in `rows` are hidden. Every other row in the span stays visible. */
  static hidden(from: number, to: number, rows: readonly number[]): RowMask {
    return new RowMask('hidden', from, to, rows);
  }

  /** Every row in the span is hidden except `rows`. */
  static shown(from: number, to: number, rows: readonly number[]): RowMask {
    return new RowMask('shown', from, to, rows);
  }

  /** Exception rows stored for this mask. The hidden count can be much larger. */
  get exceptions(): number {
    return this.#rows.length;
  }

  has(row: number): boolean {
    if (this.mode === 'none' || row < this.from || row > this.to) return false;
    const present = at(this.#rows, row);
    return this.mode === 'hidden' ? present : !present;
  }

  get size(): number {
    if (this.mode === 'none' || this.from > this.to) return 0;
    const span = this.to - this.from + 1;
    return this.mode === 'hidden' ? this.#rows.length : span - this.#rows.length;
  }

  /** Hidden rows inside the inclusive interval. */
  countIn(start: number, end: number): number {
    const from = Math.max(start, this.from);
    const to = Math.min(end, this.to);
    if (this.mode === 'none' || from > to) return 0;
    const inside = upperBound(this.#rows, to) - lowerBound(this.#rows, from);
    return this.mode === 'hidden' ? inside : (to - from + 1) - inside;
  }

  /**
   * Hidden rows in the interval that are not already accounted for by `except`.
   * Callers use this when a dimension record has already adjusted that index.
   */
  countHiddenExcept(start: number, end: number, except: ReadonlySet<number>): number {
    const from = Math.max(start, this.from);
    const to = Math.min(end, this.to);
    if (this.mode === 'none' || from > to) return 0;
    if (this.mode === 'hidden') {
      let count = 0;
      const begin = lowerBound(this.#rows, from);
      const finish = upperBound(this.#rows, to);
      for (let index = begin; index < finish; index++) if (!except.has(this.#rows[index]!)) count++;
      return count;
    }
    let count = (to - from + 1) - (upperBound(this.#rows, to) - lowerBound(this.#rows, from));
    for (const row of except) {
      if (row >= from && row <= to && this.has(row)) count--;
    }
    return count;
  }

  /** Inclusive hidden runs. A complement of a few shown rows is a few holes, not one entry per row. */
  spans(): readonly RowSpan[] {
    if (this.mode === 'none' || this.from > this.to) return [];
    if (this.mode === 'hidden') {
      const spans: RowSpan[] = [];
      for (const row of this.#rows) {
        const previous = spans.at(-1);
        if (previous && previous.end + 1 === row) spans[spans.length - 1] = { start: previous.start, end: row };
        else spans.push({ start: row, end: row });
      }
      return spans;
    }
    const spans: RowSpan[] = [];
    let cursor = this.from;
    for (const shown of this.#rows) {
      if (shown > cursor) spans.push({ start: cursor, end: shown - 1 });
      cursor = shown + 1;
    }
    if (cursor <= this.to) spans.push({ start: cursor, end: this.to });
    return spans;
  }

  /** First row at or after `start` that this mask does not hide. May fall past `to`. */
  rowAtOrAfter(start: number): number {
    if (this.mode === 'none' || start > this.to || start < this.from) return start;
    if (!this.has(start)) return start;
    if (this.mode === 'hidden') {
      let index = lowerBound(this.#rows, start);
      let end = this.#rows[index]!;
      while (index + 1 < this.#rows.length && this.#rows[index + 1] === end + 1) end = this.#rows[++index]!;
      return end + 1;
    }
    const index = lowerBound(this.#rows, start);
    return index < this.#rows.length ? this.#rows[index]! : this.to + 1;
  }

  /** First row at or before `start` that this mask does not hide. May fall before `from`. */
  rowAtOrBefore(start: number): number {
    if (this.mode === 'none' || start < this.from || start > this.to) return start;
    if (!this.has(start)) return start;
    if (this.mode === 'hidden') {
      let index = lowerBound(this.#rows, start);
      let begin = this.#rows[index]!;
      while (index > 0 && this.#rows[index - 1] === begin - 1) begin = this.#rows[--index]!;
      return begin - 1;
    }
    const index = upperBound(this.#rows, start) - 1;
    return index >= 0 ? this.#rows[index]! : this.from - 1;
  }

  equals(other: RowMask): boolean {
    if (this === other) return true;
    return this.mode === other.mode && this.from === other.from && this.to === other.to &&
      this.#rows.length === other.#rows.length &&
      this.#rows.every((value, index) => value === other.#rows[index]);
  }
}

function lowerBound(rows: readonly number[], target: number): number {
  let low = 0;
  let high = rows.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (rows[middle]! < target) low = middle + 1;
    else high = middle;
  }
  return low;
}

function upperBound(rows: readonly number[], target: number): number {
  let low = 0;
  let high = rows.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (rows[middle]! <= target) low = middle + 1;
    else high = middle;
  }
  return low;
}

function at(rows: readonly number[], target: number): boolean {
  const index = lowerBound(rows, target);
  return index < rows.length && rows[index] === target;
}

/**
 * Builds the smaller mask. A blank key that fails any criterion hides the empty run;
 * a blank key that passes every criterion hides only the failing rows.
 */
export function maskRows(from: number, to: number, criteria: readonly MaskCriterion[], blank: string,
  cells: readonly MaskCell[]): RowMask {
  if (from > to || criteria.length === 0) return RowMask.none();
  const byRow = new Map<number, Map<number, string>>();
  for (const cell of cells) {
    if (cell.row < from || cell.row > to) continue;
    let keys = byRow.get(cell.row);
    if (!keys) byRow.set(cell.row, keys = new Map());
    keys.set(cell.column, cell.key);
  }
  const passes = (keys: Map<number, string> | undefined) =>
    criteria.every(criterion => criterion.allowed.has(keys?.get(criterion.column) ?? blank));
  const blankHidden = criteria.some(criterion => !criterion.allowed.has(blank));
  const matched: number[] = [];
  for (const [row, keys] of byRow) {
    const accepted = passes(keys);
    if (blankHidden ? accepted : !accepted) matched.push(row);
  }
  matched.sort((left, right) => left - right);
  return blankHidden ? RowMask.shown(from, to, matched) : RowMask.hidden(from, to, matched);
}
