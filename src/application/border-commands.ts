import { cellBorders, draftEdges, setBorder, withBorders, type BorderDraft, type BorderEdge,
  type BorderPreset, type BorderSegment, type CellBorders, type DraftEdge } from '../core/formats/borders';
import { cellKey, documentLimits as limits, findCell, inRange, type CellBorder, type CellExtent, type CellStyle, type SheetDocument } from '../core/workbook';
import { equalDto } from '../core/workbook-delta';
import { requireDocument } from '../core/workbook-validation';
import { checkedSheet, expandMerges, type EditRange } from './editing-commands';
import type { WorkbookSnapshot } from './workbook-protocol';
import type { CellEdit, SetCellsCommand } from './workbook-session';

export function mergeAt(sheet: SheetDocument, row: number, column: number): CellExtent | undefined {
  return sheet.merges.find(merge => row >= merge.row && row < merge.row + merge.rows &&
    column >= merge.column && column < merge.column + merge.columns);
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

function sharedBorderCommand(sheet: SheetDocument, lines: Iterable<BorderEdit>, label: string,
  includeSide?: (row: number, column: number) => boolean): SetCellsCommand {
  const touched = new Map<number, { row: number; column: number; edges: CellBorders }>();

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
      const edges = cellBorders(findCell(sheet, { row, column })?.style);
      if (equalDto(edges[edge], pen ?? undefined)) return;
      entry = { row, column, edges: { ...edges } };
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
  for (const { row, column, edges } of touched.values()) {
    const target = { sheetId: sheet.id, row, column };
    const cell = findCell(sheet, target);
    if (equalDto(edges, cellBorders(cell?.style))) continue;
    const style = withBorders(cell?.style, edges);
    edits.push({ target, content: cell?.content ?? { kind: 'blank' }, preserveStyle: false,
      ...(Object.keys(style).length && { style }) });
  }
  return { kind: 'set-cells', label, edits };
}

export function applyBorderPreset(basis: WorkbookSnapshot, range: EditRange, preset: BorderPreset, pen: CellBorder | null,
  color?: string): SetCellsCommand {
  const sheet = checkedSheet(basis, range);
  const area = expandMerges(range, sheet);
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

export interface BorderEditorState {
  readonly rows: number;
  readonly columns: number;
  readonly edges: Readonly<Record<DraftEdge, CellBorder | null | 'mixed'>>;
}

export function inspectBorders(basis: WorkbookSnapshot, selected: EditRange): BorderEditorState {
  const sheet = checkedSheet(basis, selected);
  const range = expandMerges(selected, sheet);
  const merge = mergeAt(sheet, range.row, range.column);
  const collapsed = merge?.rows === range.rows && merge?.columns === range.columns;
  const edges = {} as Record<DraftEdge, CellBorder | null | 'mixed'>;
  const counts = new Map<DraftEdge, number>();
  for (const key of draftEdges) { edges[key] = null; counts.set(key, 0); }
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
  for (const key of draftEdges) {
    const { rows, columns } = range;
    const expected = key === 'horizontal' ? 2 * columns * (rows - 1) : key === 'vertical' ? 2 * rows * (columns - 1) :
      key === 'top' || key === 'bottom' ? columns : key === 'left' || key === 'right' ? rows : rows * columns;
    if (counts.get(key)! < expected && edges[key]) edges[key] = 'mixed';
  }
  return { rows: collapsed ? 1 : Math.min(range.rows, 2), columns: collapsed ? 1 : Math.min(range.columns, 2), edges };
}
