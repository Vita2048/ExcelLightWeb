import { fillPatterns, fillShadings, isFontFamily } from './formats/appearance';
import {
  cellKey, documentLimits as limits, DocumentError, formulaErrors, inRange, isCoordinate,
  type AxisDimension, type CellCache, type CellContent, type CellExtent, type CellReference,
  type CellStyle, type DocumentId, type SheetDocument, type WorkbookDocument,
} from './workbook';

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
    requireDocument(Object.keys(value).length === 2 && Object.hasOwn(value, 'kind') && Object.hasOwn(value, 'value'),
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
    case 'python':
      validateId(value.blockId);
      return requireDocument(!cache && validText(value.source, limits.pythonSourceBytes) &&
        utf8Length(value.source) <= limits.pythonSourceBytes, 'Invalid Python source (limit 1 MiB).');
    default:
      throw new DocumentError('Unsupported cell value.');
  }
}

const borderStyles = new Set(['hair', 'dotted', 'dashDotDot', 'dashDot', 'dashed', 'thin', 'mediumDashDotDot',
  'slantDashDot', 'mediumDashDot', 'mediumDashed', 'medium', 'thick', 'double']);
const borderEdges = new Set(['top', 'right', 'bottom', 'left', 'diagonalUp', 'diagonalDown']);


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
      requireDocument(borderEdges.has(side) && edge && borderStyles.has(edge.style) && validColor(edge.rgb),
        'Invalid border detail.');
    }
    requireDocument(!diagonalUp || !diagonalDown ||
      diagonalUp.style === diagonalDown.style && diagonalUp.rgb === diagonalDown.rgb, 'Diagonal borders share a style and color.');
  }
}

function validateExtent(extent: CellExtent): void {
  requireDocument(extent && isCoordinate(extent) &&
    inRange(extent.rows, 1, limits.rows - extent.row + 1) &&
    inRange(extent.columns, 1, limits.columns - extent.column + 1), 'Range is outside the sheet.');
}

function overlaps(a: CellExtent, b: CellExtent): boolean {
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

    // Mark only covered coordinates, leaving merge anchors free to hold authored values.
    let coverage: Uint8Array | undefined;
    for (const merge of sheet.merges) {
      validateExtent(merge);
      requireDocument(merge.rows * merge.columns >= 2, 'A merge requires at least two cells.');
      coverage ??= new Uint8Array(limits.rows * limits.columns);
      for (let row = merge.row; row < merge.row + merge.rows; row++) {
        for (let column = merge.column; column < merge.column + merge.columns; column++) {
          const key = cellKey({ row, column });
          requireDocument(!coverage[key], 'Merged ranges overlap.');
          coverage[key] = row === merge.row && column === merge.column ? 1 : 2;
        }
      }
    }

    const addresses = new Set<number>();
    for (const cell of sheet.cells) {
      requireDocument(cell && isCoordinate(cell), 'Cell is outside the sheet.');
      const key = cellKey(cell);
      requireDocument(!addresses.has(key), 'Duplicate cell coordinate.');
      addresses.add(key);
      addId(cell.id);
      validateCellValue(cell.content);
      if (cell.content.kind === 'formula') {
        requireDocument(++formulas <= limits.formulas, 'A workbook may contain at most 10,000 formulas.');
      }
      if (cell.content.kind === 'python') {
        requireDocument(!blocks.has(cell.content.blockId), 'Duplicate Python block identity.');
        blocks.add(cell.content.blockId);
      }
      const merged = coverage?.[key] ?? 0;
      requireDocument(!(merged && cell.content.kind === 'python') && (merged !== 2 || cell.content.kind === 'blank'),
        'Cannot put this cell in a merged range.');
      requireDocument(cell.requiresExternalCalculation === undefined || cell.requiresExternalCalculation === false ||
        cell.requiresExternalCalculation === true && cell.content.kind === 'formula', 'External calculation requires a formula.');
      requireDocument(cell.content.kind !== 'blank' || cell.style !== undefined || cell.cache !== undefined,
        'Unstyled blank cells without caches must be omitted.');
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
}
