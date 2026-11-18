import { cellKey, documentLimits, findCell, type CellExtent, type CellStyle, type IdFactory, type SheetDocument } from '../core/workbook';
import { requireDocument } from '../core/workbook-validation';
import { checkedSheet, expandMerges, type EditRange } from './editing-commands';
import { changeStyles } from './style-commands';
import type { WorkbookSnapshot } from './workbook-protocol';
import type { StructureCommand } from './workbook-session';

export const mergeLabels = { center: 'Merge & Center', across: 'Merge Across', cells: 'Merge Cells', unmerge: 'Unmerge Cells' } as const;
export type MergeMode = keyof typeof mergeLabels;
type AlignmentField = 'alignment' | 'verticalAlignment' | 'indent' | 'wrap' |
  'shrinkToFit' | 'justifyLastLine' | 'readingOrder' | 'textRotation';
export type AlignmentPatch = { readonly [K in AlignmentField]?: CellStyle[K] | undefined };
export type AlignmentAction = { kind: 'horizontal' | 'vertical'; value: string } |
  { kind: 'wrap' } | { kind: 'rotation'; value: number } | { kind: 'indent'; value: number };

export function alignmentStyle(style: CellStyle | undefined, patch: AlignmentPatch): CellStyle {
  const aligned = { ...style, ...patch };
  return Object.fromEntries(Object.entries(aligned).filter(([, value]) => value !== undefined)) as CellStyle;
}

export function changeAlignment(basis: WorkbookSnapshot, ranges: readonly EditRange[], action: AlignmentAction, active?: CellStyle) {
  return changeStyles(basis, ranges, 'Alignment', style => {
    switch (action.kind) {
      case 'horizontal': return alignmentStyle(style, {
        alignment: active?.alignment === action.value ? undefined : action.value,
        ...((action.value === 'left' || action.value === 'right') && { indent: undefined }),
      });
      case 'vertical': return alignmentStyle(style, { verticalAlignment: action.value === 'bottom' ? undefined : action.value });
      case 'wrap': return alignmentStyle(style, { wrap: !active?.wrap || undefined });
      case 'rotation': return alignmentStyle(style, { textRotation: active?.textRotation === action.value ? undefined : action.value });
      case 'indent': return alignmentStyle(style, {
        alignment: style?.alignment && ['left', 'right', 'distributed'].includes(style.alignment) ? style.alignment : 'left',
        indent: Math.max(0, Math.min(250, (style?.indent ?? 0) + action.value)) || undefined,
      });
    }
  });
}

export function overlapsMerge(a: CellExtent, b: CellExtent): boolean {
  return a.row < b.row + b.rows && b.row < a.row + a.rows && a.column < b.column + b.columns && b.column < a.column + a.columns;
}

export function mergeState(sheet: SheetDocument, area: CellExtent): boolean | undefined {
  const inside = sheet.merges.filter(merge => overlapsMerge(merge, area));
  if (inside.length === 0) return false;
  return inside.length === 1 && inside[0]!.row === area.row && inside[0]!.column === area.column &&
    inside[0]!.rows === area.rows && inside[0]!.columns === area.columns ? true : undefined;
}

export function proposeMerge(basis: WorkbookSnapshot, selection: EditRange, mode: MergeMode, makeId: IdFactory):
  { readonly command: StructureCommand; readonly discards: boolean; readonly range: EditRange } {
  requireDocument(Object.hasOwn(mergeLabels, mode), 'Unknown merge command.');
  const sheet = checkedSheet(basis, selection);
  const range = expandMerges(selection, sheet);
  const { row, column, rows, columns } = range;
  requireDocument(mode === 'unmerge' || rows * columns <= documentLimits.cells, 'Merge selection is too large.');
  const prior = sheet.merges.filter(merge => !overlapsMerge(merge, range));
  const targets: CellExtent[] = mode === 'across' && columns > 1
    ? Array.from({ length: rows }, (_, index) => ({ row: row + index, column, rows: 1, columns }))
    : mode !== 'across' && mode !== 'unmerge' && rows * columns > 1 ? [{ row, column, rows, columns }] : [];
  for (const name of basis.document.names) {
    if (name.isTable && !name.area.invalid && name.area.sheetId === sheet.id) {
      requireDocument(!targets.some(target => overlapsMerge(target, name.area)), 'Cells in a table cannot be merged.');
    }
  }
  const cells = new Map(sheet.cells.map(cell => [cellKey(cell), cell]));
  const selected = sheet.cells.filter(cell => cell.row >= row && cell.row < row + rows && cell.column >= column && cell.column < column + columns);
  const groups = new Map<number, typeof selected>();
  for (const cell of selected) {
    const key = mode === 'across' ? cell.row : row;
    const group = groups.get(key);
    if (group) group.push(cell);
    else groups.set(key, [cell]);
  }
  let discards = false;
  for (const target of targets) {
    const contents = groups.get(target.row) ?? [];
    requireDocument(!contents.some(cell => cell.content.kind === 'python'), 'Python cells and spill results cannot be merged.');
    const populated = contents.filter(cell => cell.content.kind !== 'blank' && !(cell.content.kind === 'text' && !cell.content.value));
    populated.sort((a, b) => a.row - b.row || a.column - b.column);
    discards ||= populated.length > 1;
    const key = cellKey(target);
    const source = populated[0];
    if (source && cellKey(source) !== key) {
      const anchor = cells.get(key);
      const style = anchor?.style ?? source.style;
      const numberFormat = style?.numberFormat ?? source.style?.numberFormat;
      const nextStyle = numberFormat ? { ...style, numberFormat } : style;
      const { content, cache, requiresExternalCalculation } = source;
      cells.set(key, { id: anchor?.id ?? makeId(), row: target.row, column: target.column, content,
        ...(nextStyle && { style: nextStyle }), ...(cache && { cache }), ...(requiresExternalCalculation && { requiresExternalCalculation }) });
    }
    for (const cell of contents) {
      const id = cellKey(cell);
      if (id === key) continue;
      if (cell.style) cells.set(id, { id: cell.id, row: cell.row, column: cell.column, style: cell.style, content: { kind: 'blank' } });
      else cells.delete(id);
    }
  }
  if (mode === 'center') {
    const anchor = findCell(sheet, range);
    const key = cellKey(range);
    const cell = cells.get(key) ?? anchor ?? { id: makeId(), row, column, content: { kind: 'blank' } as const };
    cells.set(key, { ...cell, style: { ...cell.style, alignment: 'center' } });
  }
  const merges = [...prior, ...targets].sort((a, b) => a.row - b.row || a.column - b.column);
  const next = { ...sheet, cells: [...cells.values()], merges };
  return { discards, range, command: { kind: 'structure', label: mergeLabels[mode],
    document: { ...basis.document, sheets: basis.document.sheets.map(item => item.id === sheet.id ? next : item) } } };
}
