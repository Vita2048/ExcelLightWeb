import { cellBorders, draftEdges, setBorder, withBorders, type BorderDraft, type BorderEdge,
  type BorderPreset, type BorderSegment, type CellBorders, type DraftEdge } from '../core/formats/borders';
import { cellKey, documentLimits as limits, findCell, findColumnStyle, findRowStyle, inRange, resolvedCellStyle,
  type AxisDimension, type CellBorder, type CellDocument, type CellExtent, type CellStyle, type SheetDocument } from '../core/workbook';
import { equalDto } from '../core/workbook-delta';
import { requireDocument } from '../core/workbook-validation';
import { checkedSheet, expandMerges, rangeSpanKind, type EditRange } from './editing-commands';
import type { WorkbookSnapshot } from './workbook-protocol';
import type { CellEdit, SetCellsCommand, StructureCommand, WorkbookCommand } from './workbook-session';

const mergesByRow = new WeakMap<SheetDocument, ReadonlyMap<number, readonly CellExtent[]>>();

export function mergeAt(sheet: SheetDocument, row: number, column: number): CellExtent | undefined {
  if (!sheet.merges.length) return undefined;
  let byRow = mergesByRow.get(sheet);
  if (!byRow) {
    const indexed = new Map<number, CellExtent[]>();
    for (const merge of sheet.merges) {
      for (let position = merge.row; position < merge.row + merge.rows; position++) {
        const entries = indexed.get(position);
        if (entries) entries.push(merge);
        else indexed.set(position, [merge]);
      }
    }
    for (const entries of indexed.values()) entries.sort((a, b) => a.column - b.column);
    byRow = indexed;
    mergesByRow.set(sheet, byRow);
  }
  const entries = byRow.get(row);
  if (!entries) return undefined;
  let low = 0, high = entries.length;
  while (low < high) {
    const index = (low + high) >>> 1;
    const merge = entries[index]!;
    if (column < merge.column) high = index;
    else if (column >= merge.column + merge.columns) low = index + 1;
    else return merge;
  }
  return undefined;
}

export interface BorderEdit extends BorderSegment { readonly pen: CellBorder | null }

export function* presetLines(range: CellExtent, preset: BorderPreset, pen: CellBorder | null, color = pen?.rgb): Generator<BorderEdit> {
  const { row, column, rows, columns } = range;
  const fixed = (style: string) => ({ style, ...(color && { rgb: color }) });
  const all = preset === 'all' || preset === 'none';
  const combined = preset === 'topBottom' || preset === 'topThickBottom' || preset === 'topDoubleBottom';
  const outline = all || preset === 'outside' || preset === 'thickOutside';
  const ink = preset === 'none' ? null : preset === 'thickOutside' ? fixed('thick') : pen;
  const bottomInk = preset === 'bottomDouble' || preset === 'topDoubleBottom' ? fixed('double') :
    preset === 'bottomThick' || preset === 'topThickBottom' ? fixed('thick') : preset === 'topBottom' ? fixed('thin') : ink;
  for (let c = column; c < column + columns; c++) {
    if (outline || preset === 'top' || combined) yield { edge: 'top', row, column: c, pen: combined ? fixed('thin') : ink };
    if (outline || preset === 'bottom' || preset === 'bottomDouble' || preset === 'bottomThick' || combined) {
      yield { edge: 'top', row: row + rows, column: c, pen: bottomInk };
    }
    if (all) for (let r = row + 1; r < row + rows; r++) yield { edge: 'top', row: r, column: c, pen: ink };
  }
  for (let r = row; r < row + rows; r++) {
    if (outline || preset === 'left') yield { edge: 'left', row: r, column, pen: ink };
    if (outline || preset === 'right') yield { edge: 'left', row: r, column: column + columns, pen: ink };
    if (all) for (let c = column + 1; c < column + columns; c++) yield { edge: 'left', row: r, column: c, pen: ink };
    if (preset === 'none') for (let c = column; c < column + columns; c++) {
      yield { edge: 'diagonalUp', row: r, column: c, pen: null };
      yield { edge: 'diagonalDown', row: r, column: c, pen: null };
    }
  }
}

function normalizeBorderStyle(style: CellStyle | undefined): CellStyle | undefined {
  return style && Object.keys(style).length ? style : undefined;
}

function reconcileBorderStyle(style: CellStyle | undefined, inherited: CellStyle | undefined): CellStyle | undefined {
  const normalized = normalizeBorderStyle(style);
  if (normalized) return equalDto(normalized, inherited) ? undefined : normalized;
  return inherited ? {} : undefined;
}

function sharedBorderCommand(sheet: SheetDocument, lines: Iterable<BorderEdit>, label: string,
  includeSide?: (row: number, column: number) => boolean): SetCellsCommand {
  const touched = new Map<number, { row: number; column: number; base: CellStyle | undefined; edges: CellBorders }>();

  function side(row: number, column: number, edge: BorderEdge, pen: CellBorder | null) {
    if (!inRange(row, 1, limits.rows) || !inRange(column, 1, limits.columns)) return;
    const merge = mergeAt(sheet, row, column);
    if (merge) {
      if (edge === 'top' && row !== merge.row || edge === 'bottom' && row !== merge.row + merge.rows - 1 ||
          edge === 'left' && column !== merge.column || edge === 'right' && column !== merge.column + merge.columns - 1) return;
      ({ row, column } = merge);
    }
    if (includeSide && !includeSide(row, column)) return;
    const key = cellKey({ row, column });
    let entry = touched.get(key);
    if (!entry) {
      const base = resolvedCellStyle(sheet, { row, column });
      const edges = cellBorders(base);
      if (equalDto(edges[edge], pen ?? undefined)) return;
      entry = { row, column, base, edges: { ...edges } };
      touched.set(key, entry);
      requireDocument(touched.size <= limits.cells, 'Border operation is too large.');
    }
    setBorder(entry.edges, edge, pen);
  }

  for (const { edge, row, column, pen } of lines) {
    requireDocument((edge === 'top' || edge === 'left' || edge === 'diagonalUp' || edge === 'diagonalDown') &&
      inRange(row, 1, limits.rows + Number(edge === 'top')) &&
      inRange(column, 1, limits.columns + Number(edge === 'left')), 'Invalid border segment.');
    side(row, column, edge, pen);
    if (edge === 'top') side(row - 1, column, 'bottom', pen);
    if (edge === 'left') side(row, column - 1, 'right', pen);
  }

  const edits: CellEdit[] = [];
  for (const { row, column, base, edges } of touched.values()) {
    const target = { sheetId: sheet.id, row, column };
    const cell = findCell(sheet, target);
    if (equalDto(edges, cellBorders(base))) continue;
    const updated = withBorders(base, edges);
    const inherited = findRowStyle(sheet, row) ?? findColumnStyle(sheet, column);
    const style = reconcileBorderStyle(updated, inherited);
    if (equalDto(style, cell?.style)) continue;
    edits.push({ target, content: cell?.content ?? { kind: 'blank' }, preserveStyle: false,
      ...(style !== undefined && { style }) });
  }
  return { kind: 'set-cells', label, edits };
}

type AxisBorderOps = ReadonlyMap<number, readonly (readonly [BorderEdge, CellBorder | null])[]>;

function applyEdgeChanges(style: CellStyle | undefined, changes: readonly (readonly [BorderEdge, CellBorder | null])[]): CellStyle | undefined {
  const edges: CellBorders = { ...cellBorders(style) };
  for (const [edge, pen] of changes) setBorder(edges, edge, pen);
  if (equalDto(edges, cellBorders(style))) return style;
  return normalizeBorderStyle(withBorders(style, edges));
}

function addAxisOp(ops: Map<number, (readonly [BorderEdge, CellBorder | null])[]>, index: number, edge: BorderEdge, pen: CellBorder | null) {
  const list = ops.get(index) ?? [];
  list.push([edge, pen]);
  ops.set(index, list);
}

export function presetBorderCommand(basis: WorkbookSnapshot, sheet: SheetDocument, area: EditRange,
  preset: BorderPreset, pen: CellBorder | null, color = pen?.rgb): StructureCommand {
  const span = rangeSpanKind(area);
  const fixed = (style: string): CellBorder => ({ style, ...(color && { rgb: color }) });
  const ink = preset === 'none' ? null : preset === 'thickOutside' ? fixed('thick') : pen;
  const bottomInk = preset === 'bottomDouble' || preset === 'topDoubleBottom' ? fixed('double') :
    preset === 'bottomThick' || preset === 'topThickBottom' ? fixed('thick') : preset === 'topBottom' ? fixed('thin') : ink;
  const topInk = preset === 'topBottom' || preset === 'topThickBottom' || preset === 'topDoubleBottom' ? fixed('thin') : ink;

  const rowOps = new Map<number, (readonly [BorderEdge, CellBorder | null])[]>();
  const colOps = new Map<number, (readonly [BorderEdge, CellBorder | null])[]>();
  const cellEdges: BorderEdit[] = [];

  if (span === 'columns') {
    const c1 = area.column, c2 = area.column + area.columns - 1;
    const all = preset === 'all' || preset === 'none';
    const outline = all || preset === 'outside' || preset === 'thickOutside';
    if (outline || preset === 'left') {
      addAxisOp(colOps, c1, 'left', ink);
      if (c1 > 1) addAxisOp(colOps, c1 - 1, 'right', ink);
    }
    if (outline || preset === 'right') {
      addAxisOp(colOps, c2, 'right', ink);
      if (c2 < limits.columns) addAxisOp(colOps, c2 + 1, 'left', ink);
    }
    if (all) {
      for (let c = c1; c <= c2; c++) {
        addAxisOp(colOps, c, 'top', ink);
        addAxisOp(colOps, c, 'bottom', ink);
        if (c < c2) addAxisOp(colOps, c, 'right', ink);
        if (c > c1) addAxisOp(colOps, c, 'left', ink);
        if (preset === 'none') {
          addAxisOp(colOps, c, 'diagonalUp', null);
          addAxisOp(colOps, c, 'diagonalDown', null);
        }
      }
    }
    const combined = preset === 'topBottom' || preset === 'topThickBottom' || preset === 'topDoubleBottom';
    if (preset === 'outside' || preset === 'thickOutside' || preset === 'top' || combined) {
      for (let c = c1; c <= c2; c++) cellEdges.push({ edge: 'top', row: 1, column: c, pen: topInk });
    }
    if (preset === 'outside' || preset === 'thickOutside' || preset === 'bottom' ||
        preset === 'bottomDouble' || preset === 'bottomThick' || combined) {
      for (let c = c1; c <= c2; c++) cellEdges.push({ edge: 'top', row: limits.rows + 1, column: c, pen: bottomInk });
    }
  } else if (span === 'rows') {
    const r1 = area.row, r2 = area.row + area.rows - 1;
    const all = preset === 'all' || preset === 'none';
    const outline = all || preset === 'outside' || preset === 'thickOutside';
    const combined = preset === 'topBottom' || preset === 'topThickBottom' || preset === 'topDoubleBottom';
    if (outline || preset === 'top' || combined) {
      addAxisOp(rowOps, r1, 'top', topInk);
      if (r1 > 1) addAxisOp(rowOps, r1 - 1, 'bottom', topInk);
    }
    if (outline || preset === 'bottom' || preset === 'bottomDouble' || preset === 'bottomThick' || combined) {
      addAxisOp(rowOps, r2, 'bottom', bottomInk);
      if (r2 < limits.rows) addAxisOp(rowOps, r2 + 1, 'top', bottomInk);
    }
    if (all) {
      for (let r = r1; r <= r2; r++) {
        addAxisOp(rowOps, r, 'left', ink);
        addAxisOp(rowOps, r, 'right', ink);
        if (r < r2) addAxisOp(rowOps, r, 'bottom', ink);
        if (r > r1) addAxisOp(rowOps, r, 'top', ink);
        if (preset === 'none') {
          addAxisOp(rowOps, r, 'diagonalUp', null);
          addAxisOp(rowOps, r, 'diagonalDown', null);
        }
      }
    }
    if (preset === 'outside' || preset === 'thickOutside' || preset === 'left') {
      requireDocument(area.rows <= limits.cells, 'Border selection is too large.');
      for (let r = r1; r <= r2; r++) cellEdges.push({ edge: 'left', row: r, column: 1, pen: ink });
    }
    if (preset === 'outside' || preset === 'thickOutside' || preset === 'right') {
      requireDocument(area.rows <= limits.cells, 'Border selection is too large.');
      for (let r = r1; r <= r2; r++) cellEdges.push({ edge: 'left', row: r, column: limits.columns + 1, pen: ink });
    }
  } else {
    // Whole sheet
    requireDocument(preset === 'all' || preset === 'none' || preset === 'top' || preset === 'bottom' ||
      preset === 'bottomDouble' || preset === 'bottomThick' || preset === 'topBottom' ||
      preset === 'topThickBottom' || preset === 'topDoubleBottom', 'Border selection is too large.');
    if (preset === 'all' || preset === 'none') {
      for (let c = 1; c <= limits.columns; c++) {
        addAxisOp(colOps, c, 'top', ink);
        addAxisOp(colOps, c, 'bottom', ink);
        addAxisOp(colOps, c, 'left', ink);
        addAxisOp(colOps, c, 'right', ink);
        if (preset === 'none') {
          addAxisOp(colOps, c, 'diagonalUp', null);
          addAxisOp(colOps, c, 'diagonalDown', null);
        }
      }
      for (const dim of sheet.rows) {
        if (!dim.style) continue;
        addAxisOp(rowOps, dim.index, 'top', ink);
        addAxisOp(rowOps, dim.index, 'bottom', ink);
        addAxisOp(rowOps, dim.index, 'left', ink);
        addAxisOp(rowOps, dim.index, 'right', ink);
        if (preset === 'none') {
          addAxisOp(rowOps, dim.index, 'diagonalUp', null);
          addAxisOp(rowOps, dim.index, 'diagonalDown', null);
        }
      }
    } else {
      const combined = preset === 'topBottom' || preset === 'topThickBottom' || preset === 'topDoubleBottom';
      if (preset === 'top' || combined) {
        for (let c = 1; c <= limits.columns; c++) cellEdges.push({ edge: 'top', row: 1, column: c, pen: topInk });
      }
      if (preset === 'bottom' || preset === 'bottomDouble' || preset === 'bottomThick' || combined) {
        for (let c = 1; c <= limits.columns; c++) cellEdges.push({ edge: 'top', row: limits.rows + 1, column: c, pen: bottomInk });
      }
    }
  }

  return applyAxisBorderChanges(basis, sheet, rowOps, colOps, cellEdges, 'Borders: ' + preset);
}

export function applyAxisBorderChanges(
  basis: WorkbookSnapshot,
  sheet: SheetDocument,
  rowOps: AxisBorderOps,
  colOps: AxisBorderOps,
  cellEdges: readonly BorderEdit[],
  label: string,
): StructureCommand {
  const rowDims = new Map(sheet.rows.map(dim => [dim.index, dim]));
  const colDims = new Map(sheet.columns.map(dim => [dim.index, dim]));
  const pending = new Map<number, { row: number; column: number; prior: CellDocument | undefined; base: CellStyle | undefined; edges: CellBorders }>();

  function getEntry(row: number, column: number) {
    const key = cellKey({ row, column });
    let entry = pending.get(key);
    if (!entry) {
      const prior = findCell(sheet, { row, column });
      const base = resolvedCellStyle(sheet, { row, column }, prior);
      const edges: CellBorders = { ...cellBorders(base) };
      const cCh = colOps.get(column);
      if (cCh) for (const [edge, pen] of cCh) setBorder(edges, edge, pen);
      const rCh = rowOps.get(row);
      if (rCh) for (const [edge, pen] of rCh) setBorder(edges, edge, pen);
      entry = { row, column, prior, base, edges };
      pending.set(key, entry);
      requireDocument(pending.size <= limits.cells, 'Border operation is too large.');
    }
    return entry;
  }

  for (const cell of sheet.cells) {
    if (!rowOps.has(cell.row) && !colOps.has(cell.column)) continue;
    getEntry(cell.row, cell.column);
  }
  for (const col of colOps.keys()) {
    for (const [row, dim] of rowDims) {
      if (!dim.style || rowOps.has(row)) continue;
      getEntry(row, col);
    }
  }
  for (const row of rowOps.keys()) {
    for (const [col, dim] of colDims) {
      if (!dim.style) continue;
      getEntry(row, col);
    }
  }

  for (const { edge, row, column, pen } of cellEdges) {
    const apply = (r: number, c: number, side: BorderEdge) => {
      if (!inRange(r, 1, limits.rows) || !inRange(c, 1, limits.columns)) return;
      const merge = mergeAt(sheet, r, c);
      if (merge) {
        if (side === 'top' && r !== merge.row || side === 'bottom' && r !== merge.row + merge.rows - 1 ||
            side === 'left' && c !== merge.column || side === 'right' && c !== merge.column + merge.columns - 1) return;
        ({ row: r, column: c } = merge);
      }
      const entry = getEntry(r, c);
      setBorder(entry.edges, side, pen);
    };
    apply(row, column, edge);
    if (edge === 'top') apply(row - 1, column, 'bottom');
    if (edge === 'left') apply(row, column - 1, 'right');
  }

  for (const [row, changes] of rowOps) {
    const dim = rowDims.get(row);
    const nextStyle = applyEdgeChanges(dim?.style, changes);
    const nextDim: AxisDimension = { index: row, ...(dim?.size !== undefined && { size: dim.size }), ...(dim?.hidden && { hidden: true }), ...(nextStyle && { style: nextStyle }) };
    if (nextDim.size === undefined && !nextDim.hidden && !nextDim.style) rowDims.delete(row);
    else rowDims.set(row, nextDim);
  }
  for (const [col, changes] of colOps) {
    const dim = colDims.get(col);
    const nextStyle = applyEdgeChanges(dim?.style, changes);
    const nextDim: AxisDimension = { index: col, ...(dim?.size !== undefined && { size: dim.size }), ...(dim?.hidden && { hidden: true }), ...(nextStyle && { style: nextStyle }) };
    if (nextDim.size === undefined && !nextDim.hidden && !nextDim.style) colDims.delete(col);
    else colDims.set(col, nextDim);
  }

  const edits: CellEdit[] = [];
  for (const { row, column, prior, base, edges } of pending.values()) {
    const updated = equalDto(edges, cellBorders(base)) ? base : withBorders(base, edges);
    const inherited = rowDims.get(row)?.style ?? colDims.get(column)?.style;
    const style = reconcileBorderStyle(updated, inherited);
    if (equalDto(style, prior?.style)) continue;
    edits.push({ target: { sheetId: sheet.id, row, column }, content: prior?.content ?? { kind: 'blank' },
      preserveStyle: false, ...(style !== undefined && { style }) });
  }

  const nextSheet: SheetDocument = {
    ...sheet,
    rows: [...rowDims.values()].sort((a, b) => a.index - b.index),
    columns: [...colDims.values()].sort((a, b) => a.index - b.index),
  };
  return {
    kind: 'structure',
    label,
    document: { ...basis.document, sheets: basis.document.sheets.map(s => s.id === sheet.id ? nextSheet : s) },
    edits,
  };
}

export function applyBorderPreset(basis: WorkbookSnapshot, range: EditRange, preset: BorderPreset, pen: CellBorder | null,
  color?: string): WorkbookCommand {
  const sheet = checkedSheet(basis, range);
  const area = expandMerges(range, sheet);
  if (rangeSpanKind(area) !== 'cells') return presetBorderCommand(basis, sheet, area, preset, pen, color);
  requireDocument(area.rows * area.columns <= limits.cells, 'Border selection is too large.');
  return sharedBorderCommand(sheet, presetLines(area, preset, pen, color), 'Borders: ' + preset);
}

export function applyBorderStroke(basis: WorkbookSnapshot, sheetId: string, segments: readonly BorderSegment[],
  pen: CellBorder | null): SetCellsCommand {
  const sheet = basis.document.sheets.find(sheet => sheet.id === sheetId);
  requireDocument(sheet && segments.length <= 26_000, 'Border stroke is too large or outside the worksheet.');
  return sharedBorderCommand(sheet, segments.map(segment => ({ ...segment, pen })), pen ? 'Draw borders' : 'Erase borders');
}

export function mirrorBorderDraft(sheet: SheetDocument, areas: readonly EditRange[], draft: BorderDraft): readonly CellEdit[] {
  const presets = [['top', draft.top], ['bottom', draft.bottom], ['left', draft.left], ['right', draft.right]] as const;
  function* perimeter() {
    for (const area of areas) {
      for (const [side, pen] of presets) {
        if (pen !== undefined) yield* presetLines(area, side, pen);
      }
    }
  }
  // Perimeter faces inside the selection are already included in the Format Cells edit.
  return sharedBorderCommand(sheet, perimeter(), 'Format Cells', (row, column) =>
    !areas.some(area => row >= area.row && row < area.row + area.rows &&
      column >= area.column && column < area.column + area.columns)).edits;
}

function sidesFor(range: CellExtent, row: number, column: number, lastRow: number, lastColumn: number, key: DraftEdge): BorderEdge[] {
  const rowEnd = range.row + range.rows - 1;
  const columnEnd = range.column + range.columns - 1;
  switch (key) {
    case 'top': return row === range.row ? ['top'] : [];
    case 'bottom': return lastRow === rowEnd ? ['bottom'] : [];
    case 'left': return column === range.column ? ['left'] : [];
    case 'right': return lastColumn === columnEnd ? ['right'] : [];
    case 'horizontal': return [...(row > range.row ? ['top'] as const : []), ...(lastRow < rowEnd ? ['bottom'] as const : [])];
    case 'vertical': return [...(column > range.column ? ['left'] as const : []), ...(lastColumn < columnEnd ? ['right'] as const : [])];
    default: return [key];
  }
}

export function applyBorderDraft(sheet: SheetDocument, range: EditRange, style: CellStyle | undefined,
  row: number, column: number, draft: BorderDraft): CellStyle | undefined {
  const merge = mergeAt(sheet, row, column);
  if (merge && (row !== merge.row || column !== merge.column)) return style;
  const edges = { ...cellBorders(style) };
  for (const [key, pen] of Object.entries(draft) as [DraftEdge, CellBorder | null][]) {
    for (const side of sidesFor(range, row, column, row + (merge?.rows ?? 1) - 1, column + (merge?.columns ?? 1) - 1, key)) {
      setBorder(edges, side, pen);
    }
  }
  return equalDto(edges, cellBorders(style)) ? style : withBorders(style, edges);
}

export function draftBorderOps(sheet: SheetDocument, areas: readonly EditRange[], draft: BorderDraft): {
  readonly rowOps: Map<number, (readonly [BorderEdge, CellBorder | null])[]>;
  readonly colOps: Map<number, (readonly [BorderEdge, CellBorder | null])[]>;
  readonly cellEdges: BorderEdit[];
} {
  const rowOps = new Map<number, (readonly [BorderEdge, CellBorder | null])[]>();
  const colOps = new Map<number, (readonly [BorderEdge, CellBorder | null])[]>();
  const cellEdges: BorderEdit[] = [];
  for (const area of areas) {
    const span = rangeSpanKind(area);
    if (span === 'cells') {
      for (const [key, pen] of Object.entries(draft) as [DraftEdge, CellBorder | null][]) {
        const { row, column, rows, columns } = area;
        if (key === 'top' || key === 'horizontal') {
          for (let c = column; c < column + columns; c++) {
            if (key === 'top') cellEdges.push({ edge: 'top', row, column: c, pen });
            else for (let r = row + 1; r < row + rows; r++) cellEdges.push({ edge: 'top', row: r, column: c, pen });
          }
        }
        if (key === 'bottom') {
          for (let c = column; c < column + columns; c++) cellEdges.push({ edge: 'top', row: row + rows, column: c, pen });
        }
        if (key === 'left' || key === 'vertical') {
          for (let r = row; r < row + rows; r++) {
            if (key === 'left') cellEdges.push({ edge: 'left', row: r, column, pen });
            else for (let c = column + 1; c < column + columns; c++) cellEdges.push({ edge: 'left', row: r, column: c, pen });
          }
        }
        if (key === 'right') {
          for (let r = row; r < row + rows; r++) cellEdges.push({ edge: 'left', row: r, column: column + columns, pen });
        }
        if (key === 'diagonalUp' || key === 'diagonalDown') {
          for (let r = row; r < row + rows; r++) {
            for (let c = column; c < column + columns; c++) cellEdges.push({ edge: key, row: r, column: c, pen });
          }
        }
      }
      continue;
    }
    if (span === 'columns') {
      const c1 = area.column, c2 = area.column + area.columns - 1;
      if (draft.left !== undefined) {
        addAxisOp(colOps, c1, 'left', draft.left);
        if (c1 > 1) addAxisOp(colOps, c1 - 1, 'right', draft.left);
      }
      if (draft.right !== undefined) {
        addAxisOp(colOps, c2, 'right', draft.right);
        if (c2 < limits.columns) addAxisOp(colOps, c2 + 1, 'left', draft.right);
      }
      if (draft.vertical !== undefined && c2 > c1) {
        for (let c = c1; c < c2; c++) addAxisOp(colOps, c, 'right', draft.vertical);
        for (let c = c1 + 1; c <= c2; c++) addAxisOp(colOps, c, 'left', draft.vertical);
      }
      if (draft.horizontal !== undefined) {
        for (let c = c1; c <= c2; c++) {
          addAxisOp(colOps, c, 'top', draft.horizontal);
          addAxisOp(colOps, c, 'bottom', draft.horizontal);
        }
      }
      if (draft.diagonalUp !== undefined) for (let c = c1; c <= c2; c++) addAxisOp(colOps, c, 'diagonalUp', draft.diagonalUp);
      if (draft.diagonalDown !== undefined) for (let c = c1; c <= c2; c++) addAxisOp(colOps, c, 'diagonalDown', draft.diagonalDown);
      if (draft.top !== undefined) for (let c = c1; c <= c2; c++) cellEdges.push({ edge: 'top', row: 1, column: c, pen: draft.top });
      if (draft.bottom !== undefined) for (let c = c1; c <= c2; c++) cellEdges.push({ edge: 'top', row: limits.rows + 1, column: c, pen: draft.bottom });
    } else if (span === 'rows') {
      const r1 = area.row, r2 = area.row + area.rows - 1;
      if (draft.top !== undefined) {
        addAxisOp(rowOps, r1, 'top', draft.top);
        if (r1 > 1) addAxisOp(rowOps, r1 - 1, 'bottom', draft.top);
      }
      if (draft.bottom !== undefined) {
        addAxisOp(rowOps, r2, 'bottom', draft.bottom);
        if (r2 < limits.rows) addAxisOp(rowOps, r2 + 1, 'top', draft.bottom);
      }
      if (draft.horizontal !== undefined && r2 > r1) {
        for (let r = r1; r < r2; r++) addAxisOp(rowOps, r, 'bottom', draft.horizontal);
        for (let r = r1 + 1; r <= r2; r++) addAxisOp(rowOps, r, 'top', draft.horizontal);
      }
      if (draft.vertical !== undefined) {
        for (let r = r1; r <= r2; r++) {
          addAxisOp(rowOps, r, 'left', draft.vertical);
          addAxisOp(rowOps, r, 'right', draft.vertical);
        }
      }
      if (draft.diagonalUp !== undefined) for (let r = r1; r <= r2; r++) addAxisOp(rowOps, r, 'diagonalUp', draft.diagonalUp);
      if (draft.diagonalDown !== undefined) for (let r = r1; r <= r2; r++) addAxisOp(rowOps, r, 'diagonalDown', draft.diagonalDown);
      if (draft.left !== undefined) {
        requireDocument(area.rows <= limits.cells, 'Border selection is too large.');
        for (let r = r1; r <= r2; r++) cellEdges.push({ edge: 'left', row: r, column: 1, pen: draft.left });
      }
      if (draft.right !== undefined) {
        requireDocument(area.rows <= limits.cells, 'Border selection is too large.');
        for (let r = r1; r <= r2; r++) cellEdges.push({ edge: 'left', row: r, column: limits.columns + 1, pen: draft.right });
      }
    } else {
      requireDocument(draft.left === undefined && draft.right === undefined, 'Border selection is too large.');
      if (draft.top !== undefined) for (let c = 1; c <= limits.columns; c++) cellEdges.push({ edge: 'top', row: 1, column: c, pen: draft.top });
      if (draft.bottom !== undefined) for (let c = 1; c <= limits.columns; c++) cellEdges.push({ edge: 'top', row: limits.rows + 1, column: c, pen: draft.bottom });
      for (let c = 1; c <= limits.columns; c++) {
        if (draft.vertical !== undefined) {
          if (c < limits.columns) addAxisOp(colOps, c, 'right', draft.vertical);
          if (c > 1) addAxisOp(colOps, c, 'left', draft.vertical);
        }
        if (draft.horizontal !== undefined) {
          addAxisOp(colOps, c, 'top', draft.horizontal);
          addAxisOp(colOps, c, 'bottom', draft.horizontal);
        }
        if (draft.diagonalUp !== undefined) addAxisOp(colOps, c, 'diagonalUp', draft.diagonalUp);
        if (draft.diagonalDown !== undefined) addAxisOp(colOps, c, 'diagonalDown', draft.diagonalDown);
      }
      for (const dim of sheet.rows) {
        if (!dim.style) continue;
        if (draft.vertical !== undefined) {
          addAxisOp(rowOps, dim.index, 'left', draft.vertical);
          addAxisOp(rowOps, dim.index, 'right', draft.vertical);
        }
        if (draft.horizontal !== undefined) {
          if (dim.index < limits.rows) addAxisOp(rowOps, dim.index, 'bottom', draft.horizontal);
          if (dim.index > 1) addAxisOp(rowOps, dim.index, 'top', draft.horizontal);
        }
        if (draft.diagonalUp !== undefined) addAxisOp(rowOps, dim.index, 'diagonalUp', draft.diagonalUp);
        if (draft.diagonalDown !== undefined) addAxisOp(rowOps, dim.index, 'diagonalDown', draft.diagonalDown);
      }
    }
  }
  return { rowOps, colOps, cellEdges };
}

export interface BorderEditorState {
  readonly rows: number;
  readonly columns: number;
  readonly edges: Readonly<Record<DraftEdge, CellBorder | null | 'mixed'>>;
}

function inspectAxisLine(
  sheet: SheetDocument,
  fixedAxis: 'row' | 'column',
  index: number,
  side: BorderEdge,
): CellBorder | null | 'mixed' {
  const axisStyle = fixedAxis === 'column' ? findColumnStyle(sheet, index) : findRowStyle(sheet, index);
  const orthoList = fixedAxis === 'column' ? sheet.rows : sheet.columns;
  const orthoCount = fixedAxis === 'column' ? limits.rows : limits.columns;
  const basePen = cellBorders(axisStyle)[side] ?? null;
  let explicitCount = 0;
  const explicitKeys = new Set<number>();
  for (const cell of sheet.cells) {
    if ((fixedAxis === 'column' ? cell.column : cell.row) !== index) continue;
    if (cell.style === undefined) continue;
    explicitKeys.add(fixedAxis === 'column' ? cell.row : cell.column);
    explicitCount++;
    const pen = cellBorders(cell.style)[side] ?? null;
    if (!equalDto(pen, basePen)) return 'mixed';
  }
  if (fixedAxis === 'column' || !axisStyle) {
    for (const dim of orthoList) {
      if (!dim.style || explicitKeys.has(dim.index)) continue;
      explicitCount++;
      const pen = cellBorders(dim.style)[side] ?? null;
      if (!equalDto(pen, basePen)) return 'mixed';
    }
  }
  // If basePen was uniform over the remaining cells, return it (if explicitCount === orthoCount, explicit items matched too).
  void orthoCount;
  void explicitCount;
  return basePen;
}

function mergeBorderReading(current: CellBorder | null | 'mixed' | undefined, next: CellBorder | null | 'mixed'): CellBorder | null | 'mixed' {
  if (current === undefined) return next;
  if (current === 'mixed' || next === 'mixed') return 'mixed';
  return equalDto(current, next) ? current : 'mixed';
}

export function inspectBorders(basis: WorkbookSnapshot, selected: EditRange): BorderEditorState {
  const sheet = checkedSheet(basis, selected);
  const range = expandMerges(selected, sheet);
  const span = rangeSpanKind(range);
  if (span !== 'cells') {
    const edges = {} as Record<DraftEdge, CellBorder | null | 'mixed'>;
    if (span === 'columns') {
      const c1 = range.column, c2 = range.column + range.columns - 1;
      const sampleEdge = (row: number, side: BorderEdge) => {
        let acc: CellBorder | null | 'mixed' | undefined;
        for (let c = c1; c <= c2; c++) acc = mergeBorderReading(acc, cellBorders(resolvedCellStyle(sheet, { row, column: c }))[side] ?? null);
        return acc ?? null;
      };
      edges.top = sampleEdge(1, 'top');
      edges.bottom = sampleEdge(limits.rows, 'bottom');
      edges.left = inspectAxisLine(sheet, 'column', c1, 'left');
      edges.right = inspectAxisLine(sheet, 'column', c2, 'right');
      let v: CellBorder | null | 'mixed' | undefined;
      for (let c = c1; c < c2; c++) v = mergeBorderReading(v, inspectAxisLine(sheet, 'column', c, 'right'));
      for (let c = c1 + 1; c <= c2; c++) v = mergeBorderReading(v, inspectAxisLine(sheet, 'column', c, 'left'));
      edges.vertical = c2 > c1 ? (v ?? null) : null;
      let h: CellBorder | null | 'mixed' | undefined;
      for (let c = c1; c <= c2; c++) {
        h = mergeBorderReading(h, inspectAxisLine(sheet, 'column', c, 'top'));
        h = mergeBorderReading(h, inspectAxisLine(sheet, 'column', c, 'bottom'));
      }
      edges.horizontal = h ?? null;
      let du: CellBorder | null | 'mixed' | undefined, dd: CellBorder | null | 'mixed' | undefined;
      for (let c = c1; c <= c2; c++) {
        du = mergeBorderReading(du, inspectAxisLine(sheet, 'column', c, 'diagonalUp'));
        dd = mergeBorderReading(dd, inspectAxisLine(sheet, 'column', c, 'diagonalDown'));
      }
      edges.diagonalUp = du ?? null;
      edges.diagonalDown = dd ?? null;
    } else if (span === 'rows') {
      const r1 = range.row, r2 = range.row + range.rows - 1;
      const sampleEdge = (column: number, side: BorderEdge) => {
        let acc: CellBorder | null | 'mixed' | undefined;
        for (let r = r1; r <= r2; r++) acc = mergeBorderReading(acc, cellBorders(resolvedCellStyle(sheet, { row: r, column }))[side] ?? null);
        return acc ?? null;
      };
      edges.left = sampleEdge(1, 'left');
      edges.right = sampleEdge(limits.columns, 'right');
      edges.top = inspectAxisLine(sheet, 'row', r1, 'top');
      edges.bottom = inspectAxisLine(sheet, 'row', r2, 'bottom');
      let h: CellBorder | null | 'mixed' | undefined;
      for (let r = r1; r < r2; r++) h = mergeBorderReading(h, inspectAxisLine(sheet, 'row', r, 'bottom'));
      for (let r = r1 + 1; r <= r2; r++) h = mergeBorderReading(h, inspectAxisLine(sheet, 'row', r, 'top'));
      edges.horizontal = h ?? null;
      let v: CellBorder | null | 'mixed' | undefined;
      for (let r = r1; r <= r2; r++) {
        v = mergeBorderReading(v, inspectAxisLine(sheet, 'row', r, 'left'));
        v = mergeBorderReading(v, inspectAxisLine(sheet, 'row', r, 'right'));
      }
      edges.vertical = v ?? null;
      let du: CellBorder | null | 'mixed' | undefined, dd: CellBorder | null | 'mixed' | undefined;
      for (let r = r1; r <= r2; r++) {
        du = mergeBorderReading(du, inspectAxisLine(sheet, 'row', r, 'diagonalUp'));
        dd = mergeBorderReading(dd, inspectAxisLine(sheet, 'row', r, 'diagonalDown'));
      }
      edges.diagonalUp = du ?? null;
      edges.diagonalDown = dd ?? null;
    } else {
      for (const key of draftEdges) edges[key] = null;
      const sampleAll = (side: BorderEdge) => {
        let acc: CellBorder | null | 'mixed' | undefined;
        for (let c = 1; c <= limits.columns; c++) {
          acc = mergeBorderReading(acc, cellBorders(findColumnStyle(sheet, c))[side] ?? null);
          if (acc === 'mixed') return 'mixed' as const;
        }
        return acc ?? null;
      };
      edges.top = sampleAll('top');
      edges.bottom = sampleAll('bottom');
      edges.left = sampleAll('left');
      edges.right = sampleAll('right');
      edges.vertical = mergeBorderReading(edges.left, edges.right);
      edges.horizontal = mergeBorderReading(edges.top, edges.bottom);
      edges.diagonalUp = sampleAll('diagonalUp');
      edges.diagonalDown = sampleAll('diagonalDown');
    }
    return { rows: 2, columns: 2, edges };
  }

  const merge = mergeAt(sheet, range.row, range.column);
  const collapsed = merge?.rows === range.rows && merge?.columns === range.columns;
  const edges = {} as Record<DraftEdge, CellBorder | null | 'mixed'>;
  const counts = new Map<DraftEdge, number>();
  for (const key of draftEdges) { edges[key] = null; counts.set(key, 0); }
  const hasAxisBorders = (sheet.rows.some(r => r.style && r.index >= range.row && r.index < range.row + range.rows &&
      (r.style.borders || r.style.borderDetails)) ||
    sheet.columns.some(c => c.style && c.index >= range.column && c.index < range.column + range.columns &&
      (c.style.borders || c.style.borderDetails)));
  if (hasAxisBorders && range.rows * range.columns <= 10_000) {
    for (let r = range.row; r < range.row + range.rows; r++) {
      for (let c = range.column; c < range.column + range.columns; c++) {
        const span = mergeAt(sheet, r, c);
        if (span && (span.row !== r || span.column !== c)) continue;
        const border = cellBorders(resolvedCellStyle(sheet, { row: r, column: c }));
        for (const key of draftEdges) {
          const sides = sidesFor(range, r, c, r + (span?.rows ?? 1) - 1, c + (span?.columns ?? 1) - 1, key);
          for (const side of sides) {
            const prior = counts.get(key)!;
            const next = border[side] ?? null;
            if (!prior) edges[key] = next;
            else if (!equalDto(edges[key], next)) edges[key] = 'mixed';
            const weight = side === 'top' || side === 'bottom' ? span?.columns ?? 1 :
              side === 'left' || side === 'right' ? span?.rows ?? 1 : (span?.rows ?? 1) * (span?.columns ?? 1);
            counts.set(key, prior + weight);
          }
        }
      }
    }
  } else {
    for (const cell of sheet.cells) {
      if (cell.row < range.row || cell.column < range.column || cell.row >= range.row + range.rows ||
          cell.column >= range.column + range.columns) continue;
      const span = mergeAt(sheet, cell.row, cell.column);
      if (span && (span.row !== cell.row || span.column !== cell.column)) continue;
      const border = cellBorders(cell.style);
      for (const key of draftEdges) {
        const sides = sidesFor(range, cell.row, cell.column, cell.row + (span?.rows ?? 1) - 1,
          cell.column + (span?.columns ?? 1) - 1, key);
        for (const side of sides) {
          const prior = counts.get(key)!;
          const next = border[side] ?? null;
          if (!prior) edges[key] = next;
          else if (!equalDto(edges[key], next)) edges[key] = 'mixed';
          const weight = side === 'top' || side === 'bottom' ? span?.columns ?? 1 :
            side === 'left' || side === 'right' ? span?.rows ?? 1 : (span?.rows ?? 1) * (span?.columns ?? 1);
          counts.set(key, prior + weight);
        }
      }
    }
  }
  for (const key of draftEdges) {
    const { rows, columns } = range;
    const expected = key === 'horizontal' ? 2 * columns * (rows - 1) : key === 'vertical' ? 2 * rows * (columns - 1) :
      key === 'top' || key === 'bottom' ? columns : key === 'left' || key === 'right' ? rows : rows * columns;
    if (counts.get(key)! < expected && edges[key]) edges[key] = 'mixed';
  }
  return { rows: collapsed ? 1 : Math.min(range.rows, 2), columns: collapsed ? 1 : Math.min(range.columns, 2), edges };
}
