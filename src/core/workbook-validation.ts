import { isValidChartType, type ChartSeries } from './chart';
import {
  pivotAggregations, pivotReportLayouts, pivotShowValuesAsModes, pivotStyles, pivotSubtotalPositions,
} from './pivot';
import { borderStyles, borderEdges } from './formats/borders';
import { fillPatterns, fillShadings, isFontFamily } from './formats/appearance';
import {
  cellKey, documentLimits as limits, DocumentError, formulaErrors, inRange, isCoordinate, worksheetViewport,
  type AxisDimension, type CellCache, type CellContent, type CellDocument, type CellExtent, type CellReference,
  type CellStyle, type Coordinate, type DocumentId, type FigurePlacement, type PythonFigure, type SheetDocument, type WorkbookDocument,
} from './workbook';

const pythonBindingReference = /^=(?:(?:'(?:[^']|'')+'|[^!']+)!)?\$?[A-Za-z]+\$?[0-9]+(?::\$?[A-Za-z]+\$?[0-9]+)?$/u;

export function requireDocument(condition: unknown, message: string): asserts condition {
  if (!condition) throw new DocumentError(message);
}

export function validateId(id: DocumentId): void {
  requireDocument(typeof id === 'string' && id.trim().length > 0 && id.length <= 128 &&
    !/[\u0000-\u001f]/u.test(id) && id !== '00000000-0000-0000-0000-000000000000', 'Invalid document identity.');
}

export function validateTarget(target: CellReference, sheetIds: ReadonlySet<DocumentId>): void {
  requireDocument(!target.invalid && isCoordinate(target) && sheetIds.has(target.sheetId),
    'Cell reference is invalid or outside workbook limits.');
}

function validText(value: string, limit: number = limits.textLength): boolean {
  return typeof value === 'string' && value.length <= limit && !value.includes('\0');
}

export function ordinalNameKey(value: string): string {
  let key = '';
  for (const char of value) {
    const upper = char.toUpperCase();
    // Desktop ordinal comparisons use one-to-one case mappings and keep
    // the ASCII and non-ASCII ranges distinct.
    key += upper.length === char.length && (char <= '\x7f' || upper > '\x7f') ? upper : char;
  }
  return key;
}

// Count UTF-8 bytes without depending on a host encoder.
function utf8Length(source: string): number {
  let bytes = 0;
  for (const char of source) {
    const point = char.codePointAt(0)!;
    bytes += point <= 0x7f ? 1 : point <= 0x7ff ? 2 : point <= 0xffff ? 3 : 4;
  }
  return bytes;
}

export function validateCellValue(value: CellContent | CellCache, cache = false): void {
  requireDocument(value && typeof value === 'object', 'Cell value is missing.');
  if (cache) {
    const keys = Object.keys(value);
    const hasSerializedText = Object.hasOwn(value, 'serializedText');
    const serializedText = (value as { readonly serializedText?: unknown }).serializedText;
    requireDocument(Object.hasOwn(value, 'kind') && Object.hasOwn(value, 'value') &&
      keys.every(key => key === 'kind' || key === 'value' || key === 'serializedText') &&
      (!hasSerializedText || value.kind === 'number' && typeof serializedText === 'string' &&
        serializedText.length <= limits.textLength &&
        /^[\t\n\v\f\r ]*[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?[\t\n\v\f\r ]*$/u.test(serializedText) &&
        Number(serializedText) === (value as { readonly value: number }).value),
    'Calculation cache must contain a kind and scalar value.');
  }
  switch (value.kind) {
    case 'number':
      return requireDocument(Number.isFinite(value.value), 'Numeric values must be finite.');
    case 'boolean':
      return requireDocument(typeof value.value === 'boolean', 'Invalid Boolean value.');
    case 'text':
      return requireDocument(validText(value.value), 'Cell text is missing, too long, or contains a NUL.');
    case 'error':
      return requireDocument(cache && formulaErrors.includes(value.value), 'Invalid calculation error.');
    case 'blank':
      return requireDocument(!cache, 'Invalid calculation cache.');
    case 'formula':
      return requireDocument(!cache && validText(value.source, limits.formulaLength) &&
        value.source.length >= 2 && value.source.startsWith('='), 'A formula must start with = and have valid source text.');
    case 'python': {
      validateId(value.blockId);
      requireDocument(!cache && validText(value.source, limits.pythonSourceBytes) &&
        utf8Length(value.source) <= limits.pythonSourceBytes, 'Invalid Python source (limit 1 MiB).');
      if (value.block) {
        requireDocument(value.block.id === value.blockId, 'Python block id must match.');
        requireDocument(value.block.source === value.source, 'Invalid Python id/source (limit 1 MiB).');
        requireDocument(Array.isArray(value.block.bindings) && value.block.bindings.length <= 100, 'Invalid or duplicate Python bindings.');
        const bindingNames = new Set<string>();
        for (const b of value.block.bindings) {
          requireDocument(b && typeof b.name === 'string' && /^[A-Za-z_][A-Za-z_0-9]{0,63}$/.test(b.name), 'Bindings need a unique name and one A1 cell/range reference.');
          requireDocument(!bindingNames.has(b.name), 'Invalid or duplicate Python bindings.');
          bindingNames.add(b.name);
          requireDocument(typeof b.reference === 'string' && b.reference.length <= 512 &&
            (b.reference === '=#REF!' || pythonBindingReference.test(b.reference)),
          'Bindings need a unique name and one A1 cell/range reference.');
        }
        if (value.block.status !== undefined) {
          requireDocument(['never', 'succeeded', 'failed', 'cancelled', 'timedOut', 'workerExited'].includes(value.block.status), 'Invalid Python status.');
        }
        if (value.block.error != null) {
          requireDocument(typeof value.block.error === 'string' && value.block.error.length <= 32767, 'Python metadata exceeds limits.');
        }
        if (value.block.inputFingerprint != null) {
          requireDocument(typeof value.block.inputFingerprint === 'string' && value.block.inputFingerprint.length <= 128, 'Python metadata exceeds limits.');
        }
        if (value.block.runtimeFingerprint != null) {
          requireDocument(typeof value.block.runtimeFingerprint === 'string' && value.block.runtimeFingerprint.length <= 128, 'Python metadata exceeds limits.');
        }
        if (value.block.lastValue != null) {
          validateCellValue(value.block.lastValue, true);
          requireDocument(value.block.hasResult, 'Python value requires result metadata.');
        }
        if (value.block.lastFigure != null) {
          requireDocument(value.block.hasResult && value.block.lastTable == null && value.block.lastValue == null, 'Invalid figure result metadata.');
          validateFigure(value.block.lastFigure);
        }
        if (value.block.lastTable != null) {
          const t = value.block.lastTable;
          requireDocument(inRange(t.rows, 1, limits.rows) && inRange(t.columns, 1, limits.columns) && t.rows * t.columns <= 100000, 'Invalid table shape/notices (limit 100,000 cells).');
          requireDocument(Array.isArray(t.cells) && t.cells.length === t.rows * t.columns, 'Invalid table shape/notices (limit 100,000 cells).');
          for (const cache of t.cells) {
            if (cache === null) continue;
            requireDocument(cache.kind !== 'error', 'Unsupported table value.');
            validateCellValue(cache, true);
          }
          requireDocument(Array.isArray(t.notices) && t.notices.length <= 100 && t.notices.every(n => typeof n === 'string' && n.length <= 32767), 'Invalid table shape/notices (limit 100,000 cells).');
          requireDocument(value.block.hasResult, 'Invalid Python table result metadata.');
          requireDocument(equalCache(t.cells[0] ?? null, value.block.lastValue ?? null), 'Python table anchor differs from its cached first value.');
        }
        if (value.block.spillActive) {
          requireDocument(value.block.lastTable != null, 'Invalid Python table result metadata.');
        }
        if (value.block.placements) {
          requireDocument(Array.isArray(value.block.placements) && value.block.placements.length <= 32, 'Invalid figure placements.');
        }
      }
      return;
    }
    default:
      throw new DocumentError('Unsupported cell value.');
  }
}

function equalCache(a: CellCache | null, b: CellCache | null): boolean {
  if (a === null && b === null) return true;
  if (a === null || b === null) return false;
  return a.kind === b.kind && (a as { value: unknown }).value === (b as { value: unknown }).value &&
    (a.kind !== 'number' || b.kind !== 'number' || a.serializedText === b.serializedText);
}

function validateFigure(fig: PythonFigure): void {
  requireDocument(fig && typeof fig.pngBase64 === 'string' && Number.isInteger(fig.width) && Number.isInteger(fig.height), 'Invalid figure.');
  requireDocument(fig.width >= 1 && fig.width <= 8192 && fig.height >= 1 && fig.height <= 8192 && fig.width * fig.height <= 16_000_000,
    'Figure exceeds 16 million pixels / 8192 pixels per side.');
  requireDocument(fig.pngBase64.length <= 22_369_624, 'Figure PNG exceeds 16 MiB.');
}



function validColor(value: string | undefined): boolean {
  return value === undefined || typeof value === 'string' && /^[0-9A-F]{6}$/u.test(value);
}

function validateStyle(style: CellStyle): void {
  requireDocument(style && typeof style === 'object' && !Array.isArray(style), 'Invalid cell style.');
  requireDocument(!style.doubleUnderline || style.underline, 'Double underline requires underline.');
  requireDocument(style.script === undefined || ['superscript', 'subscript'].includes(style.script), 'Invalid font script.');
  requireDocument(style.underlineKind === undefined || style.underline &&
    ['singleAccounting', 'doubleAccounting'].includes(style.underlineKind), 'Invalid underline style.');
  requireDocument(style.fontFamily === undefined || typeof style.fontFamily === 'string' && isFontFamily(style.fontFamily), 'Invalid font family.');
  requireDocument(style.fontSize === undefined || Number.isFinite(style.fontSize) &&
    style.fontSize >= 6 && style.fontSize <= 72, 'Font size must be 6 to 72 points.');
  requireDocument(style.alignment === undefined || ['left', 'center', 'right', 'fill', 'justify',
    'centerContinuous', 'distributed'].includes(style.alignment), 'Invalid horizontal alignment.');
  requireDocument(style.verticalAlignment === undefined ||
    ['top', 'center', 'bottom', 'justify', 'distributed'].includes(style.verticalAlignment), 'Invalid vertical alignment.');
  requireDocument(style.textRotation === undefined || style.textRotation === 255 ||
    inRange(style.textRotation, -90, 90), 'Invalid text rotation.');
  requireDocument(style.indent === undefined || inRange(style.indent, 0, 250), 'Invalid indent.');
  requireDocument(style.readingOrder === undefined || ['ltr', 'rtl'].includes(style.readingOrder), 'Invalid text direction.');
  requireDocument(style.borders === undefined ||
    /^[trbl]{0,4}$/u.test(style.borders) && new Set(style.borders).size === style.borders.length, 'Invalid borders.');
  requireDocument(style.numberFormat === undefined || validText(style.numberFormat, limits.numberFormatLength) &&
    style.numberFormat.length > 0 && !/[\u0000-\u001f]/u.test(style.numberFormat), 'Invalid number format.');
  requireDocument(validColor(style.textRgb) && validColor(style.fillRgb), 'Colors must be six uppercase hex digits.');
  if (style.fillDetails !== undefined) {
    const fill = style.fillDetails;
    requireDocument(fill && typeof fill.pattern === 'string' && Object.hasOwn(fillPatterns, fill.pattern) && inRange(fill.variant ?? 0, 0, 3) &&
      validColor(fill.patternRgb) && validColor(fill.gradientRgb), 'Invalid fill effect.');
    requireDocument(fill.shading === undefined || fill.pattern === 'none' && fill.gradientRgb &&
      typeof fill.shading === 'string' && Object.hasOwn(fillShadings, fill.shading), 'Invalid gradient.');
  }
  if (style.borderDetails !== undefined) {
    requireDocument(style.borderDetails && typeof style.borderDetails === 'object', 'Invalid border detail.');
    const { diagonalUp, diagonalDown } = style.borderDetails;
    for (const [side, edge] of Object.entries(style.borderDetails)) {
      requireDocument(borderEdges.some(key => key === side) && edge && typeof edge.style === 'string' &&
        edge.style !== 'none' && Object.hasOwn(borderStyles, edge.style) && validColor(edge.rgb),
        'Invalid border detail.');
    }
    requireDocument(!diagonalUp || !diagonalDown ||
      diagonalUp.style === diagonalDown.style && diagonalUp.rgb === diagonalDown.rgb, 'Diagonal borders share a style and color.');
  }
}

export function validateExtent(extent: CellExtent): void {
  requireDocument(extent && isCoordinate(extent) &&
    inRange(extent.rows, 1, limits.rows - extent.row + 1) &&
    inRange(extent.columns, 1, limits.columns - extent.column + 1), 'Range is outside the sheet.');
}

/** 1 is the merge anchor, 2 is a covered cell, 0 is outside every merge. */
function mergeRole(merges: readonly CellExtent[], cell: Coordinate): 0 | 1 | 2 {
  for (const merge of merges) {
    if (cell.row < merge.row || cell.row >= merge.row + merge.rows ||
        cell.column < merge.column || cell.column >= merge.column + merge.columns) continue;
    return cell.row === merge.row && cell.column === merge.column ? 1 : 2;
  }
  return 0;
}

export function overlaps(a: CellExtent, b: CellExtent): boolean {
  return a.row < b.row + b.rows && b.row < a.row + a.rows &&
    a.column < b.column + b.columns && b.column < a.column + a.columns;
}

function validateDimensions(dimensions: readonly AxisDimension[], max: number): void {
  requireDocument(Array.isArray(dimensions), 'Sheet dimensions are missing.');
  const indexes = new Set<number>();
  for (const dimension of dimensions) {
    requireDocument(dimension && inRange(dimension.index, 1, max) && !indexes.has(dimension.index) &&
      (dimension.size === undefined || Number.isFinite(dimension.size) && dimension.size >= 4 && dimension.size <= 4096) &&
      (dimension.hidden === undefined || typeof dimension.hidden === 'boolean'), 'Invalid or duplicate dimension.');
    if (dimension.style !== undefined) validateStyle(dimension.style);
    indexes.add(dimension.index);
  }
}

function validateFilter(filter: SheetDocument['filter']): void {
  if (filter === undefined) return;
  validateExtent(filter);
  const max = filter.column + filter.columns - 1;
  requireDocument(Array.isArray(filter.criteria) && filter.criteria.length <= filter.columns, 'Invalid filter criteria.');
  requireDocument(filter.sort === undefined || filter.sort && inRange(filter.sort.column, filter.column, max) &&
    typeof filter.sort.ascending === 'boolean', 'Invalid filter sort.');
  const columns = new Set<number>();
  for (const criterion of filter.criteria) {
    requireDocument(criterion && inRange(criterion.column, filter.column, max) && !columns.has(criterion.column),
      'Invalid or duplicate filter column.');
    columns.add(criterion.column);
    requireDocument(Array.isArray(criterion.allowed) && criterion.allowed.length <= limits.rows &&
      criterion.allowed.every((value: string) => validText(value)) &&
      new Set(criterion.allowed).size === criterion.allowed.length, 'Invalid or duplicate filter value.');
  }
}

function validateSheets(document: WorkbookDocument, addId: (id: DocumentId) => void): void {
  const sheetNames = new Set<string>();
  const blocks = new Set<DocumentId>();
  let cells = 0;
  let formulas = 0;
  for (const sheet of document.sheets) {
    requireDocument(sheet, 'Sheet is missing.');
    addId(sheet.id);
    requireDocument(validText(sheet.name, limits.sheetNameLength) && sheet.name.trim() &&
      !/^'|'$|[\u0000-\u001f:\\/?*\[\]]/u.test(sheet.name), 'Invalid sheet name.');
    const normalizedName = ordinalNameKey(sheet.name);
    requireDocument(!sheetNames.has(normalizedName), 'Duplicate sheet name.');
    sheetNames.add(normalizedName);
    validateDimensions(sheet.rows, limits.rows);
    validateDimensions(sheet.columns, limits.columns);
    validateFilter(sheet.filter);
    requireDocument(sheet.frozenPanes && inRange(sheet.frozenPanes.rows, 0, limits.rows - 1) &&
      inRange(sheet.frozenPanes.columns, 0, limits.columns - 1), 'Invalid frozen panes.');
    requireDocument(Array.isArray(sheet.cells) && Array.isArray(sheet.merges), 'Sheet cells or merges are missing.');
    cells += sheet.cells.length;
    requireDocument(cells <= limits.cells, 'A workbook may contain at most 200,000 populated cells.');

    // Overlap is pairwise. Covered cells are found from populated records, not a grid bitmap.
    for (const merge of sheet.merges) {
      validateExtent(merge);
      requireDocument(merge.rows * merge.columns >= 2, 'A merge requires at least two cells.');
    }
    for (let index = 0; index < sheet.merges.length; index++) {
      for (let other = index + 1; other < sheet.merges.length; other++) {
        requireDocument(!overlaps(sheet.merges[index]!, sheet.merges[other]!), 'Merged ranges overlap.');
      }
    }

    const addresses = new Set<number>();
    for (const cell of sheet.cells) {
      requireDocument(cell && isCoordinate(cell), 'Cell is outside the sheet.');
      const key = cellKey(cell);
      requireDocument(!addresses.has(key), 'Duplicate cell coordinate.');
      addresses.add(key);
      addId(cell.id);
      if (cell.spillOwner !== undefined) {
        validateId(cell.spillOwner);
      }
      if (cell.spillKind !== undefined) {
        requireDocument(cell.spillKind === 'python' || cell.spillKind === 'formula' || cell.spillKind === 'pivot', 'Invalid spill kind.');
        requireDocument(cell.spillOwner !== undefined, 'Spill kind requires spill ownership.');
      }
      if (cell.spillExtent !== undefined) {
        requireDocument(cell.content.kind === 'formula' && inRange(cell.spillExtent.rows, 1, limits.rows) &&
          inRange(cell.spillExtent.columns, 1, limits.columns) && cell.spillExtent.rows * cell.spillExtent.columns <= 100_000,
          'Invalid formula spill extent.');
      }
      validateCellValue(cell.content);
      if (cell.content.kind === 'formula') {
        requireDocument(++formulas <= limits.formulas, 'A workbook may contain at most 10,000 formulas.');
      }
      if (cell.content.kind === 'python') {
        requireDocument(!blocks.has(cell.content.blockId), 'Duplicate Python block identity.');
        blocks.add(cell.content.blockId);
      }
      const merged = mergeRole(sheet.merges, cell);
      requireDocument(!(merged && cell.content.kind === 'python') && (merged !== 2 || cell.content.kind === 'blank'),
        'Cannot put this cell in a merged range.');
      requireDocument(cell.requiresExternalCalculation === undefined || cell.requiresExternalCalculation === false ||
        cell.requiresExternalCalculation === true && cell.content.kind === 'formula', 'External calculation requires a formula.');
      requireDocument(cell.content.kind !== 'blank' || cell.style !== undefined || cell.cache !== undefined || cell.spillOwner !== undefined,
        'Blank cells are omitted unless they carry a style, cache or spill ownership.');
      if (cell.style !== undefined) validateStyle(cell.style);
      if (cell.cache !== undefined) validateCellValue(cell.cache, true);
    }
  }
}

function validateNames(document: WorkbookDocument, addId: (id: DocumentId) => void): void {
  requireDocument(Array.isArray(document.names) && document.names.length <= limits.names,
    'A workbook may contain at most 1,000 names and tables.');
  const names = new Set<string>();
  const tables: WorkbookDocument['names'][number][] = [];
  for (const entry of document.names) {
    requireDocument(entry && typeof entry.name === 'string' && /^[A-Za-z_][A-Za-z_0-9]{0,254}$/u.test(entry.name) &&
      !/^(?:[a-z]+[0-9]+|r[0-9]*c[0-9]*|r|c)$/iu.test(entry.name), 'Invalid workbook name.');
    addId(entry.id);
    const key = entry.name.toUpperCase();
    requireDocument(!names.has(key), 'Duplicate workbook name.');
    names.add(key);
    requireDocument(entry.comment === undefined || validText(entry.comment), 'Invalid name comment.');
    validateExtent(entry.area);
    validateId(entry.area.sheetId);
    const sheet = document.sheets.find(sheet => sheet.id === entry.area.sheetId);
    requireDocument(entry.area.invalid || sheet, 'Name refers to a missing sheet.');
    const design = entry.design;
    requireDocument(design === undefined || design && entry.isTable &&
      ['Neutral', 'Blue', 'Green', 'Orange', 'Purple'].includes(design.style) &&
      (!design.totalRow || entry.area.rows >= (design.headerRow ? 2 : 1)), 'Invalid table design.');
    const labels = entry.headerLabels ?? [];
    requireDocument(Array.isArray(labels) && labels.length <= limits.columns && labels.every(label => validText(label)) &&
      (entry.isTable || labels.length === 0), 'Invalid detached table schema.');
    if (entry.isTable && !entry.area.invalid && sheet) {
      requireDocument(design?.headerRow !== false || labels.length === entry.area.columns &&
        labels.every(label => label.trim()) && new Set(labels.map(ordinalNameKey)).size === labels.length,
        'Table column labels must be nonblank and unique.');
      requireDocument(!tables.some(table => table.area.sheetId === sheet.id && overlaps(entry.area, table.area)),
        'Table extents cannot overlap.');
      requireDocument(!sheet.merges.some(merge => overlaps(entry.area, merge)), 'Cells in a table cannot be merged.');
      tables.push(entry);
    }
  }
}

function equalContentToCache(content: CellContent, expected: CellCache | null): boolean {
  if (expected == null) {
    return content.kind === 'blank';
  }
  switch (expected.kind) {
    case 'number':
      return content.kind === 'number' && expected.value === content.value;
    case 'boolean':
      return content.kind === 'boolean' && expected.value === content.value;
    case 'text':
      return content.kind === 'text' && expected.value === content.value;
    default:
      return content.kind === 'blank';
  }
}

function validateSpills(document: WorkbookDocument): void {
  const pythonChildren = new Map<DocumentId, { sheetId: DocumentId; cell: CellDocument }[]>();
  const formulaChildren = new Map<DocumentId, { sheetId: DocumentId; cell: CellDocument }[]>();
  const pivotChildren = new Map<DocumentId, { sheetId: DocumentId; cell: CellDocument }[]>();
  for (const sheet of document.sheets) {
    for (const cell of sheet.cells) {
      if (cell.spillOwner !== undefined) {
        const target = cell.spillKind === 'formula'
          ? formulaChildren
          : cell.spillKind === 'pivot'
            ? pivotChildren
            : pythonChildren;
        const list = target.get(cell.spillOwner) ?? [];
        list.push({ sheetId: sheet.id, cell });
        target.set(cell.spillOwner, list);
      }
    }
  }

  for (const [ownerId, children] of pythonChildren) {
    let anchorSheet: SheetDocument | undefined;
    let anchorCell: CellDocument | undefined;
    for (const s of document.sheets) {
      const match = s.cells.find(c => c.content.kind === 'python' && c.content.blockId === ownerId);
      if (match) {
        anchorSheet = s;
        anchorCell = match;
        break;
      }
    }
    if (!anchorSheet || !anchorCell || anchorCell.content.kind !== 'python') {
      throw new DocumentError('Spill anchor is missing.');
    }
    const block = anchorCell.content.block;
    if (!block || !block.spillActive || !block.lastTable) {
      throw new DocumentError('Orphaned spill children.');
    }
    const table = block.lastTable;
    const startRow = anchorCell.row;
    const startCol = anchorCell.column;
    const lastR = startRow + table.rows - 1;
    const lastC = startCol + table.columns - 1;
    requireDocument(inRange(lastR, 1, limits.rows) && inRange(lastC, 1, limits.columns), 'Selection exceeds rectangular cell limits.');

    for (const { sheetId, cell } of children) {
      if (sheetId !== anchorSheet.id || cell.row < startRow || cell.row > lastR || cell.column < startCol || cell.column > lastC || (cell.row === startRow && cell.column === startCol)) {
        throw new DocumentError('Cannot change part of a Python spill; edit or move its complete output at the anchor.');
      }
      const expectedIndex = (cell.row - startRow) * table.columns + (cell.column - startCol);
      const expected = table.cells[expectedIndex] ?? null;
      if (!equalContentToCache(cell.content, expected)) {
        throw new DocumentError(`Cannot change part of a Python spill at row ${cell.row}, column ${cell.column}; edit or move its complete output at the anchor.`);
      }
    }
  }

  for (const [ownerId, children] of formulaChildren) {
    let anchorSheet: SheetDocument | undefined;
    let anchorCell: CellDocument | undefined;
    for (const s of document.sheets) {
      const match = s.cells.find(c => c.id === ownerId && c.content.kind === 'formula');
      if (match) {
        anchorSheet = s;
        anchorCell = match;
        break;
      }
    }
    if (!anchorSheet || !anchorCell || !anchorCell.spillExtent) {
      throw new DocumentError('Formula spill anchor is missing.');
    }
    const lastR = anchorCell.row + anchorCell.spillExtent.rows - 1;
    const lastC = anchorCell.column + anchorCell.spillExtent.columns - 1;
    for (const { sheetId, cell } of children) {
      if (sheetId !== anchorSheet.id || cell.row < anchorCell.row || cell.row > lastR ||
          cell.column < anchorCell.column || cell.column > lastC ||
          (cell.row === anchorCell.row && cell.column === anchorCell.column)) {
        throw new DocumentError('Cannot change part of a formula spill; edit or move its complete output at the anchor.');
      }
    }
  }

  const pivotsById = new Map((document.pivots ?? []).map(pivot => [pivot.id, pivot]));
  for (const [ownerId, children] of pivotChildren) {
    const owner = pivotsById.get(ownerId);
    requireDocument(owner && owner.extent, 'PivotTable spill owner is missing.');
    const lastR = owner.extent.row + owner.extent.rows - 1;
    const lastC = owner.extent.column + owner.extent.columns - 1;
    for (const { sheetId, cell } of children) {
      requireDocument(sheetId === owner.target.sheetId && cell.row >= owner.extent.row && cell.row <= lastR &&
        cell.column >= owner.extent.column && cell.column <= lastC,
      'Cannot change part of a PivotTable; modify or refresh the PivotTable instead.');
    }
  }

  for (const sheet of document.sheets) {
    for (const cell of sheet.cells) {
      if (cell.content.kind === 'formula' && cell.spillExtent) {
        const startRow = cell.row;
        const startCol = cell.column;
        const lastR = startRow + cell.spillExtent.rows - 1;
        const lastC = startCol + cell.spillExtent.columns - 1;
        requireDocument(inRange(lastR, 1, limits.rows) && inRange(lastC, 1, limits.columns), 'Selection exceeds rectangular cell limits.');
        const children = formulaChildren.get(cell.id) ?? [];
        requireDocument(children.length === cell.spillExtent.rows * cell.spillExtent.columns - 1,
          'Formula spill children do not match spill extent.');
        const extent = { row: startRow, column: startCol, rows: cell.spillExtent.rows, columns: cell.spillExtent.columns };
        requireDocument(!sheet.merges.some(merge => overlaps(merge, extent)), 'Formula spill cannot overlap merged cells.');
        for (const name of document.names) {
          if (name.isTable && !name.area.invalid && name.area.sheetId === sheet.id) {
            requireDocument(!overlaps(name.area, extent), 'Formula spill cannot overlap a workbook table.');
          }
        }
        for (const pivot of document.pivots ?? []) {
          if (pivot.target.sheetId === sheet.id && pivot.extent) {
            requireDocument(!overlaps(pivot.extent, extent), 'Formula spill cannot overlap a PivotTable.');
          }
        }
      }
      if (cell.content.kind !== 'python' || !cell.content.block?.spillActive) continue;
      const table = cell.content.block.lastTable;
      if (!table) continue;
      const startRow = cell.row;
      const startCol = cell.column;
      const lastR = startRow + table.rows - 1;
      const lastC = startCol + table.columns - 1;
      requireDocument(inRange(lastR, 1, limits.rows) && inRange(lastC, 1, limits.columns), 'Selection exceeds rectangular cell limits.');
      const children = pythonChildren.get(cell.content.blockId) ?? [];
      if (children.length !== table.cells.length - 1) {
        throw new DocumentError('Cannot split a Python spill; select its complete output or its anchor.');
      }
      const pyExtent = { row: startRow, column: startCol, rows: table.rows, columns: table.columns };
      for (const name of document.names) {
        if (name.isTable && !name.area.invalid && name.area.sheetId === sheet.id) {
          if (overlaps(name.area, pyExtent)) {
            throw new DocumentError('Python spill cannot overlap a workbook table.');
          }
        }
      }
      for (const pivot of document.pivots ?? []) {
        if (pivot.target.sheetId === sheet.id && pivot.extent && overlaps(pivot.extent, pyExtent)) {
          throw new DocumentError('Python spill cannot overlap a PivotTable.');
        }
      }
    }
  }
}

function validateFigures(document: WorkbookDocument): void {
  const placements: FigurePlacement[] = [];
  for (const sheet of document.sheets) {
    for (const cell of sheet.cells) {
      if (cell.content.kind === 'python' && cell.content.block?.placements) {
        for (const p of cell.content.block.placements) {
          placements.push(p);
        }
      }
    }
  }
  if (placements.length > 32) {
    throw new DocumentError('Maximum 32 distinct figure placements per workbook.');
  }
  const placementIds = new Set<string>();
  for (const p of placements) {
    validateId(p.id);
    if (placementIds.has(p.id)) {
      throw new DocumentError('Maximum 32 distinct figure placements per workbook.');
    }
    placementIds.add(p.id);
    const sheet = document.sheets.find(s => s.id === p.sheetId);
    requireDocument(sheet, 'Placement refers to a missing sheet.');
    requireDocument(isCoordinate({ row: p.row, column: p.column }), 'Placement coordinate is outside grid.');
    requireDocument(inRange(p.rows, 1, limits.rows - p.row + 1), 'Placement row extent is invalid.');
    requireDocument(inRange(p.columns, 1, limits.columns - p.column + 1), 'Placement column extent is invalid.');
    requireDocument(p.row + p.rows - 1 <= worksheetViewport.rows && p.column + p.columns - 1 <= worksheetViewport.columns, 'Placement extent is outside viewport.');
    const { offsetX = 0, offsetY = 0, width = 0, height = 0 } = p;
    requireDocument(Number.isFinite(offsetX) && Number.isFinite(offsetY) && Number.isFinite(width) && Number.isFinite(height), 'Invalid floating figure position or size.');
    requireDocument(offsetX >= 0 && offsetX <= 4096 && offsetY >= 0 && offsetY <= 4096, 'Invalid floating figure position or size.');
    requireDocument((width === 0 && height === 0) || (width >= 80 && height >= 60 && width <= 49152 && height <= 4096000), 'Invalid floating figure position or size.');
    requireDocument((p.renderedFigure == null) === (p.renderedSourceId == null), 'Invalid chart render cache identity.');
    if (p.renderedSourceId != null) {
      requireDocument(p.renderedSourceId.length === 64 && /^[0-9a-fA-F]{64}$/.test(p.renderedSourceId), 'Invalid chart render cache identity.');
    }
    if (p.renderedFigure) {
      validateFigure(p.renderedFigure);
    }
  }
}

function validatePivots(document: WorkbookDocument, addId: (id: DocumentId) => void): void {
  if (document.pivots === undefined) return;
  requireDocument(Array.isArray(document.pivots) && document.pivots.length <= 200,
    'A workbook may contain at most 200 PivotTables.');
  const sheetIds = new Set(document.sheets.map(sheet => sheet.id));
  const pivotNames = new Set<string>();
  const placed: { readonly sheetId: DocumentId; readonly extent: CellExtent }[] = [];

  for (const pivot of document.pivots) {
    requireDocument(pivot && typeof pivot === 'object', 'Invalid PivotTable entry.');
    addId(pivot.id);
    requireDocument(typeof pivot.name === 'string' && pivot.name.trim().length > 0 && validText(pivot.name, 255),
      'PivotTable requires a valid name.');
    const nameKey = ordinalNameKey(pivot.name.trim());
    requireDocument(!pivotNames.has(nameKey), 'Duplicate PivotTable name.');
    pivotNames.add(nameKey);

    validateTarget(pivot.target, sheetIds);
    const sheet = document.sheets.find(s => s.id === pivot.target.sheetId)!;

    if (pivot.source.kind === 'range') {
      validateExtent(pivot.source.area);
      requireDocument(pivot.source.area.invalid || sheetIds.has(pivot.source.area.sheetId),
        'PivotTable source refers to a missing sheet.');
    } else {
      validateId(pivot.source.tableId);
    }

    if (pivot.extent !== undefined) {
      validateExtent(pivot.extent);
      requireDocument(pivot.target.row >= pivot.extent.row &&
        pivot.target.row < pivot.extent.row + pivot.extent.rows &&
        pivot.target.column >= pivot.extent.column &&
        pivot.target.column < pivot.extent.column + pivot.extent.columns,
      'PivotTable target must lie inside its extent.');
      requireDocument(!sheet.merges.some(merge => overlaps(merge, pivot.extent!)),
        'PivotTable cannot overlap merged cells.');
      for (const name of document.names) {
        if (name.isTable && !name.area.invalid && name.area.sheetId === sheet.id) {
          requireDocument(!overlaps(name.area, pivot.extent), 'PivotTable cannot overlap a workbook table.');
        }
      }
      requireDocument(!placed.some(other => other.sheetId === sheet.id && overlaps(other.extent, pivot.extent!)),
        'PivotTables cannot overlap each other.');
      placed.push({ sheetId: sheet.id, extent: pivot.extent });
    }

    requireDocument(pivotReportLayouts.includes(pivot.reportLayout), 'Invalid PivotTable report layout.');
    requireDocument(pivotSubtotalPositions.includes(pivot.subtotals), 'Invalid PivotTable subtotals position.');
    requireDocument(pivotStyles.includes(pivot.style), 'Invalid PivotTable style.');
    for (const af of [...pivot.rows, ...pivot.columns]) {
      requireDocument(
        af.numberFormat === undefined ||
        (typeof af.numberFormat === 'string' && validText(af.numberFormat, limits.numberFormatLength) &&
          af.numberFormat.length > 0 && !/[\u0000-\u001f]/u.test(af.numberFormat)),
        'Invalid PivotTable field number format.',
      );
    }
    for (const vf of pivot.values) {
      validateId(vf.id);
      requireDocument(pivotAggregations.includes(vf.aggregation), 'Invalid PivotTable aggregation.');
      requireDocument(vf.showValuesAs === undefined || pivotShowValuesAsModes.includes(vf.showValuesAs),
        'Invalid PivotTable Show Values As calculation.');
      requireDocument(
        vf.numberFormat === undefined ||
        (typeof vf.numberFormat === 'string' && validText(vf.numberFormat, limits.numberFormatLength) &&
          vf.numberFormat.length > 0 && !/[\u0000-\u001f]/u.test(vf.numberFormat)),
        'Invalid PivotTable value field number format.',
      );
    }
  }
}

function validateChartSeries(series: ChartSeries, seenIds: Set<DocumentId>): void {
  requireDocument(series && typeof series === 'object', 'Invalid chart series entry.');
  validateId(series.id);
  requireDocument(!seenIds.has(series.id), 'Chart series identities must be unique within a chart.');
  seenIds.add(series.id);
  requireDocument(typeof series.valuesRef === 'string' && series.valuesRef.length > 0 && validText(series.valuesRef, limits.formulaLength),
    'Chart series requires a valid values reference.');
  for (const optionalRef of [series.nameRef, series.categoriesRef, series.xValuesRef, series.bubbleSizesRef]) {
    if (optionalRef !== undefined) {
      requireDocument(typeof optionalRef === 'string' && optionalRef.length > 0 && validText(optionalRef, limits.formulaLength),
        'Invalid chart series reference.');
    }
  }
  if (series.chartType !== undefined) {
    requireDocument(isValidChartType(series.chartType), 'Invalid chart series type.');
  }
  if (series.gapWidth !== undefined) {
    requireDocument(Number.isFinite(series.gapWidth) && series.gapWidth >= 0 && series.gapWidth <= 500, 'Invalid chart gap width.');
  }
  if (series.overlap !== undefined) {
    requireDocument(Number.isFinite(series.overlap) && series.overlap >= -100 && series.overlap <= 100, 'Invalid chart overlap.');
  }
  if (series.holeSize !== undefined) {
    requireDocument(Number.isFinite(series.holeSize) && series.holeSize >= 10 && series.holeSize <= 90, 'Invalid doughnut hole size.');
  }
  if (series.firstSliceAngle !== undefined) {
    requireDocument(Number.isFinite(series.firstSliceAngle) && series.firstSliceAngle >= 0 && series.firstSliceAngle <= 360, 'Invalid first slice angle.');
  }
  if (series.explosion !== undefined) {
    requireDocument(Number.isFinite(series.explosion) && series.explosion >= 0 && series.explosion <= 100, 'Invalid slice explosion.');
  }
  if (series.lineWidth !== undefined) {
    requireDocument(Number.isFinite(series.lineWidth) && series.lineWidth >= 0 && series.lineWidth <= 20, 'Invalid series line width.');
  }
  if (series.binCount !== undefined) {
    requireDocument(inRange(series.binCount, 1, 1000), 'Invalid histogram bin count.');
  }
  if (series.binWidth !== undefined) {
    requireDocument(Number.isFinite(series.binWidth) && series.binWidth > 0, 'Invalid histogram bin width.');
  }
  if (series.trendlines !== undefined) {
    requireDocument(Array.isArray(series.trendlines) && series.trendlines.length <= 16, 'Invalid chart trendlines.');
    for (const trend of series.trendlines) {
      validateId(trend.id);
      requireDocument(['linear', 'exponential', 'logarithmic', 'polynomial', 'power', 'movingAverage'].includes(trend.type),
        'Invalid chart trendline type.');
      if (trend.polynomialOrder !== undefined) {
        requireDocument(inRange(trend.polynomialOrder, 2, 6), 'Polynomial trendline order must be 2 to 6.');
      }
      if (trend.period !== undefined) {
        requireDocument(inRange(trend.period, 2, 1000), 'Moving average trendline period must be at least 2.');
      }
    }
  }
}

function validateCharts(document: WorkbookDocument, addId: (id: DocumentId) => void): void {
  if (document.charts === undefined) return;
  requireDocument(Array.isArray(document.charts) && document.charts.length <= 200,
    'A workbook may contain at most 200 charts.');
  const sheetIds = new Set(document.sheets.map(sheet => sheet.id));
  const chartNames = new Set<string>();

  for (const chart of document.charts) {
    requireDocument(chart && typeof chart === 'object', 'Invalid chart entry.');
    addId(chart.id);
    requireDocument(typeof chart.name === 'string' && chart.name.trim().length > 0 && validText(chart.name, 255),
      'Chart requires a valid name.');
    const nameKey = ordinalNameKey(chart.name.trim());
    requireDocument(!chartNames.has(nameKey), 'Duplicate chart name.');
    chartNames.add(nameKey);

    requireDocument(isValidChartType(chart.chartType), 'Invalid chart type.');
    requireDocument(chart.seriesBy === 'columns' || chart.seriesBy === 'rows', 'Invalid chart series orientation.');
    requireDocument(chart.hiddenEmptyMode === 'gaps' || chart.hiddenEmptyMode === 'zero' || chart.hiddenEmptyMode === 'connect',
      'Invalid chart hidden/empty cell mode.');
    requireDocument(typeof chart.showHiddenCells === 'boolean', 'Invalid chart showHiddenCells flag.');
    if (chart.showFieldButtons !== undefined) {
      requireDocument(typeof chart.showFieldButtons === 'boolean', 'Invalid chart showFieldButtons flag.');
    }
    requireDocument(inRange(chart.styleId, 1, 48), 'Invalid chart style id.');
    requireDocument(typeof chart.paletteId === 'string' && chart.paletteId.length > 0 && validText(chart.paletteId, 64),
      'Invalid chart palette id.');

    const p = chart.placement;
    requireDocument(p && typeof p === 'object' && sheetIds.has(p.sheetId), 'Chart placement refers to a missing sheet.');
    requireDocument(isCoordinate({ row: p.row, column: p.column }), 'Chart anchor coordinate is outside grid.');
    requireDocument(Number.isFinite(p.offsetX) && Number.isFinite(p.offsetY) &&
      p.offsetX >= -4096 && p.offsetX <= 4096 && p.offsetY >= -4096 && p.offsetY <= 4096,
    'Invalid chart placement offset.');
    requireDocument(Number.isFinite(p.width) && p.width >= 120 && p.width <= 4096 &&
      Number.isFinite(p.height) && p.height >= 90 && p.height <= 4096,
    'Invalid chart dimensions.');
    if (p.mode !== undefined) {
      requireDocument(p.mode === 'moveAndSize' || p.mode === 'moveOnly' || p.mode === 'freeFloating',
        'Invalid chart placement mode.');
    }
    if (p.zIndex !== undefined) {
      requireDocument(Number.isInteger(p.zIndex), 'Invalid chart zIndex.');
    }

    requireDocument(chart.source && typeof chart.source === 'object', 'Invalid chart source.');
    if (chart.source.kind === 'range') {
      requireDocument(typeof chart.source.formula === 'string' && chart.source.formula.length > 0 &&
        validText(chart.source.formula, limits.formulaLength),
      'Invalid chart source formula.');
    } else if (chart.source.kind === 'table') {
      validateId(chart.source.tableId);
    } else if (chart.source.kind === 'pivot') {
      validateId(chart.source.pivotId);
    } else {
      throw new DocumentError('Unsupported chart source kind.');
    }

    requireDocument(Array.isArray(chart.series) && chart.series.length <= 255, 'Invalid chart series list.');
    const seriesIds = new Set<DocumentId>();
    for (const entry of chart.series) {
      validateChartSeries(entry, seriesIds);
    }
    requireDocument(chart.elements && typeof chart.elements === 'object', 'Invalid chart elements.');
  }
}

export function validateWorkbook(document: WorkbookDocument): void {
  requireDocument(document?.formatVersion === 1 && document.dateSystem === '1900', 'Unsupported document version or date system.');
  requireDocument(Array.isArray(document.sheets) && inRange(document.sheets.length, 1, limits.sheets),
    'A workbook requires 1 to 10 sheets.');
  const ids = new Set<DocumentId>();
  function addId(id: DocumentId): void {
    validateId(id);
    requireDocument(!ids.has(id), 'Workbook, sheet, cell and name identities must be unique.');
    ids.add(id);
  }
  addId(document.id);
  validateSheets(document, addId);
  validateNames(document, addId);
  validatePivots(document, addId);
  validateCharts(document, addId);
  validateSpills(document);
  validateFigures(document);
}
