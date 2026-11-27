import type { Element } from '@xmldom/xmldom';
import { BlobReader, ZipReader } from '@zip.js/zip.js';
import { parseFormula, copyReferences } from '../../core/formula-references';
import type {
  CellCache, CellDocument, CellExtent, CellStyle, IdFactory, SheetDocument, WorkbookDocument,
} from '../../core/workbook';
import { documentLimits, DocumentError, formulaErrors } from '../../core/workbook';
import { validateWorkbook } from '../../core/workbook-validation';
import { supportedFunctions } from '../../calculation/formula-functions';
import { desktopZipBudget } from '../bounded-zip';
import { PackageException } from '../native';
import { readStyles } from './xlsx-styles';
import {
  attr, childElements, firstChild, isTruthyAttribute, parseInvariantNumber, parseXmlPart,
  relationshipNamespace, textOf,
} from './xlsx-xml';
import { exchangeCodes, XlsxReport } from './xlsx-report';

const cfbSignature = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const worksheetPart = /^xl\/worksheets\/[^/]+\.xml$/iu;
const maxReportedNames = 12;

export interface XlsxImport {
  readonly document: WorkbookDocument;
  readonly report: XlsxReport;
}

interface PackageContent {
  readonly parts: ReadonlyMap<string, Uint8Array>;
  readonly entries: readonly string[];
}

function isWantedPart(name: string): boolean {
  const lower = name.toLowerCase();
  return lower === '[content_types].xml' || lower === 'xl/workbook.xml' ||
    lower === 'xl/_rels/workbook.xml.rels' || lower === 'xl/sharedstrings.xml' ||
    lower === 'xl/styles.xml' || worksheetPart.test(name) ||
    lower.endsWith('.rels') && lower.startsWith('xl/worksheets/_rels/');
}

/** Bounded sequential package read. Only the OOXML parts this adapter maps are retained. */
async function readPackage(blob: Blob): Promise<PackageContent> {
  if (blob.size > desktopZipBudget.compressedBytes) {
    throw new PackageException(`XLSX exceeds the ${desktopZipBudget.compressedBytes} byte limit. No document was adopted.`);
  }
  const head = new Uint8Array(await blob.slice(0, 8).arrayBuffer());
  if (cfbSignature.every((byte, index) => head[index] === byte)) {
    throw new PackageException('Encrypted/password-protected or legacy Excel files are unsupported. Save an unencrypted .xlsx file first.');
  }
  const reader = new ZipReader(new BlobReader(blob), {
    strictness: 'strict', checkCrc32: true, checkOverlappingEntry: true, useWebWorkers: false,
  });
  const parts = new Map<string, Uint8Array>();
  const entries: string[] = [];
  const seen = new Set<string>();
  let expanded = 0;
  try {
    for await (const entry of reader.getEntriesGenerator()) {
      if (entries.length >= desktopZipBudget.entries) {
        throw new PackageException('XLSX has more than ' + desktopZipBudget.entries + ' entries. No document was adopted.');
      }
      if (entry.directory) continue;
      const name = entry.filename.replaceAll('\\', '/');
      if (name.includes('..') || name.startsWith('/') || name.includes(':') ||
        name.split('/').some(segment => !segment || segment === '.')) {
        throw new PackageException(`XLSX entry path is not allowed: ${entry.filename}`);
      }
      const key = name.toLowerCase();
      if (seen.has(key)) throw new PackageException('Duplicate XLSX package entry. No document was adopted.');
      seen.add(key);
      if (entry.encrypted) throw new PackageException('Encrypted XLSX packages are unsupported. No document was adopted.');
      entries.push(name);
      if (entry.uncompressedSize > desktopZipBudget.entryBytes) {
        throw new PackageException('XLSX exceeds expanded package limits. No document was adopted.');
      }
      expanded += entry.uncompressedSize;
      if (expanded > desktopZipBudget.expandedBytes) {
        throw new PackageException('XLSX exceeds expanded package limits. No document was adopted.');
      }
      if (!isWantedPart(name)) continue;
      const chunks: Uint8Array[] = [];
      let size = 0;
      await entry.getData(new WritableStream<Uint8Array>({
        write(chunk) {
          size += chunk.byteLength;
          if (size > desktopZipBudget.entryBytes) {
            throw new PackageException('XLSX exceeds expanded package limits. No document was adopted.');
          }
          chunks.push(chunk);
        },
      }));
      const bytes = new Uint8Array(size);
      let offset = 0;
      for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.byteLength; }
      parts.set(name.toLowerCase(), bytes);
    }
  } catch (error) {
    if (error instanceof PackageException) throw error;
    throw new PackageException('XLSX could not be read as a ZIP package. No document was adopted.');
  } finally {
    await reader.close().catch(() => {});
  }
  return { parts, entries };
}

function scanEntries(entries: readonly string[], report: XlsxReport): void {
  for (const name of entries) {
    const lower = name.toLowerCase();
    if (lower === 'xl/vbaproject.bin') {
      throw new PackageException('Macro-enabled Excel files are unsupported, including renamed .xlsm files.');
    }
    if (lower.startsWith('xl/charts/') || lower.startsWith('xl/chartsheets/')) {
      report.add(exchangeCodes.charts, 'Charts were not imported.');
    } else if (lower.startsWith('xl/pivotcache/') || lower.startsWith('xl/pivottables/')) {
      report.add(exchangeCodes.pivots, 'Pivot tables were not imported.');
    } else if (lower.startsWith('xl/drawings/')) {
      report.add(exchangeCodes.drawings, 'Drawings were not imported.');
    } else if (lower.startsWith('xl/comments') || lower.startsWith('xl/threadedcomments')) {
      report.add(exchangeCodes.comments, 'Comments were not imported.');
    } else if (lower.startsWith('xl/externallinks/')) {
      report.add(exchangeCodes.externalLinks, 'External links were not imported.');
    } else if (lower.startsWith('xl/tables/')) {
      report.add(exchangeCodes.tables, 'Tables were not imported.');
    } else if (lower.startsWith('xl/media/')) {
      report.add(exchangeCodes.pictures, 'Pictures were not imported.');
    }
  }
}

function checkContentType(parts: ReadonlyMap<string, Uint8Array>): void {
  const bytes = parts.get('[content_types].xml');
  if (!bytes) return;
  const { root } = parseXmlPart(bytes, 'The XLSX content-types part');
  for (const element of childElements(root, 'Override')) {
    const partName = attr(element, 'PartName')?.toLowerCase();
    const contentType = attr(element, 'ContentType') ?? '';
    if (partName === '/xl/workbook.xml' &&
      /macroenabled|binary/i.test(contentType)) {
      throw new PackageException('Macro-enabled or binary Excel workbooks are unsupported. Save an unencrypted .xlsx file first.');
    }
  }
}

function relationshipId(element: Element): string | undefined {
  const qualified = element.getAttributeNS(relationshipNamespace, 'id');
  return qualified === null || qualified === '' ? attr(element, 'r:id') : qualified;
}

function resolveTarget(target: string): string {
  const text = target.replaceAll('\\', '/');
  if (text.includes('..')) throw new PackageException('XLSX sheet target is not allowed.');
  if (text.startsWith('/')) return text.slice(1);
  return 'xl/' + text.replace(/^\/+/u, '');
}

function workspaceSerial(iso: string): number {
  const parsed = Date.parse(iso);
  if (!Number.isFinite(parsed)) throw new PackageException('XLSX contains an invalid date value. No document was adopted.');
  return parsed / 86_400_000 + 25_569;
}

function errorValue(text: string): typeof formulaErrors[number] {
  const key = text.trim().toUpperCase();
  if (!(formulaErrors as readonly string[]).includes(key)) {
    throw new PackageException(`Unsupported error ${text}. No document was adopted.`);
  }
  return key as typeof formulaErrors[number];
}

function formulaNeedsExternalCalculation(source: string): boolean {
  let expression;
  try {
    expression = parseFormula(source).expression;
  } catch {
    return true;
  }
  function visit(node: ReturnType<typeof parseFormula>['expression']): boolean {
    switch (node.kind) {
      case 'call': return !supportedFunctions.has(node.name) || node.args.some(visit);
      case 'name':
      case 'array':
        return true;
      case 'union': return node.parts.some(visit);
      case 'unary': return visit(node.value);
      case 'binary': return visit(node.left) || visit(node.right);
      default:
        return false;
    }
  }
  return visit(expression);
}

function sharedFormulaText(master: string, dr: number, dc: number): string | undefined {
  if (dr === 0 && dc === 0) return master;
  try {
    return copyReferences('=' + master, dr, dc).slice(1);
  } catch {
    return undefined;
  }
}

interface RawCell {
  readonly row: number;
  readonly column: number;
  readonly address: string;
  readonly type?: string;
  readonly styleIndex?: number;
  readonly value?: string;
  readonly inline?: string;
  readonly formula?: { readonly text?: string; readonly kind?: string; readonly si?: string };
}

function parseArea(reference: string): CellExtent | undefined {
  const parts = reference.split(':');
  if (parts.length > 2) return undefined;
  const endpoint = (text: string) => {
    const match = /^([A-Z]{1,3})([1-9][0-9]*)$/u.exec(text.trim().toUpperCase());
    if (!match?.[1] || !match[2]) return undefined;
    const column = [...match[1]].reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0);
    return { row: Number(match[2]), column };
  };
  const first = endpoint(parts[0]!);
  const last = endpoint(parts[1] ?? parts[0]!);
  if (!first || !last) return undefined;
  return {
    row: Math.min(first.row, last.row), column: Math.min(first.column, last.column),
    rows: Math.abs(last.row - first.row) + 1, columns: Math.abs(last.column - first.column) + 1,
  };
}

function cellText(cell: Element, shared: readonly string[]): string | undefined {
  const type = attr(cell, 't');
  if (type === 'inlineStr') {
    const inline = firstChild(cell, 'is');
    return childElements(inline).map(node => textOf(node) ?? '').join('');
  }
  const raw = textOf(firstChild(cell, 'v'));
  if (raw === undefined) return undefined;
  if (type === 's') {
    const index = Number(raw);
    if (!Number.isInteger(index) || index < 0 || index >= shared.length) {
      throw new PackageException('XLSX contains an invalid shared-string reference. No document was adopted.');
    }
    return shared[index]!;
  }
  return raw;
}

function readSharedStrings(bytes: Uint8Array | undefined, report: XlsxReport): string[] {
  if (!bytes) return [];
  const { root } = parseXmlPart(bytes, 'The XLSX shared strings part');
  return childElements(root, 'si').map(item => {
    if (childElements(item, 'r').length > 0) {
      report.add(exchangeCodes.richText, 'Rich-text runs are flattened to plain text with cell formatting.');
    }
    const direct = childElements(item, 't');
    if (direct.length > 0) return direct.map(node => textOf(node) ?? '').join('');
    return childElements(item, 'r')
      .map(run => childElements(run, 't').map(node => textOf(node) ?? '').join('')).join('');
  });
}

function readWorksheet(raw: Element, sheetName: string, shared: readonly string[], styles: ReturnType<typeof readStyles>,
  use1904: boolean, report: XlsxReport, makeId: IdFactory, counters: { cells: number; formulas: number }): SheetDocument {
  const sheetData = firstChild(raw, 'sheetData');
  const rows = sheetData ? childElements(sheetData, 'row') : [];
  const rawCells: RawCell[] = [];
  const masters = new Map<string, RawCell>();
  let previousRow = 0;

  for (const row of rows) {
    const rawRow = attr(row, 'r');
    const rowIndex = rawRow === undefined ? previousRow + 1 : Number(rawRow);
    previousRow = rowIndex;
    if (!Number.isInteger(rowIndex) || rowIndex < 1 || rowIndex > documentLimits.rows) {
      throw new PackageException('Row dimensions are outside supported workbook limits. No document was adopted.');
    }
    for (const cell of childElements(row, 'c')) {
      const address = attr(cell, 'r');
      if (!address) throw new PackageException('XLSX cell address is missing. No document was adopted.');
      const parsed = parseArea(address);
      if (!parsed || parsed.rows !== 1 || parsed.columns !== 1) {
        throw new PackageException('XLSX cell is outside supported workbook limits. No document was adopted.');
      }
      if (parsed.row > documentLimits.rows || parsed.column > documentLimits.columns) {
        throw new PackageException('XLSX cell is outside supported workbook limits. No document was adopted.');
      }
      if (++counters.cells > documentLimits.cells) {
        throw new PackageException('XLSX exceeds supported cell count. No document was adopted.');
      }
      const formula = firstChild(cell, 'f');
      let formulaRecord: RawCell['formula'];
      if (formula) {
        if (++counters.formulas > documentLimits.formulas) {
          throw new PackageException('XLSX exceeds supported formula count. No document was adopted.');
        }
        formulaRecord = {
          ...(textOf(formula) ? { text: textOf(formula)! } : {}),
          ...(attr(formula, 't') ? { kind: attr(formula, 't')! } : {}),
          ...(attr(formula, 'si') ? { si: attr(formula, 'si')! } : {}),
        };
      }
      const styleIndex = attr(cell, 's') === undefined ? undefined : Number(attr(cell, 's'));
      if (styleIndex !== undefined && (!Number.isInteger(styleIndex) || styleIndex < 0)) {
        throw new PackageException('XLSX style index is invalid. No document was adopted.');
      }
      const inline = cellText(cell, shared);
      const record: RawCell = {
        row: parsed.row, column: parsed.column, address,
        ...(attr(cell, 't') ? { type: attr(cell, 't')! } : {}),
        ...(styleIndex !== undefined ? { styleIndex } : {}),
        ...(textOf(firstChild(cell, 'v')) !== undefined ? { value: textOf(firstChild(cell, 'v'))! } : {}),
        ...(attr(cell, 't') === 'inlineStr' && inline !== undefined ? { inline } : {}),
        ...(formulaRecord ? { formula: formulaRecord } : {}),
      };
      if (formulaRecord?.kind === 'shared' && formulaRecord.si && formulaRecord.text) {
        masters.set(formulaRecord.si, record);
      }
      if (formulaRecord?.kind === 'array') {
        report.add(exchangeCodes.arrayFormula, `${sheetName}!${address} is an array formula. Its text is kept and it is not evaluated.`);
      }
      rawCells.push(record);
    }
  }

  let usedMaxRow = 1;
  let usedMaxColumn = 1;
  for (const record of rawCells) {
    if (record.row > usedMaxRow) usedMaxRow = record.row;
    if (record.column > usedMaxColumn) usedMaxColumn = record.column;
  }
  const dimensions = readDimensions(raw, usedMaxRow, usedMaxColumn, styles, report);
  const styledRows = new Set(dimensions.rows.filter(r => r.style !== undefined).map(r => r.index));
  const styledColumns = new Set(dimensions.columns.filter(c => c.style !== undefined).map(c => c.index));

  const cells: CellDocument[] = [];
  for (const record of rawCells) {
    const rawStyle = styles.cellStyle(record.styleIndex);
    const hasAxisStyle = styledRows.has(record.row) || styledColumns.has(record.column);
    const style = rawStyle ?? (record.styleIndex !== undefined && hasAxisStyle ? {} : undefined);
    const effectiveStyleIndex = record.styleIndex ?? dimensions.rowStyleByIndex.get(record.row) ?? dimensions.colStyleByIndex.get(record.column);
    const dateFormatted = styles.isDateFormatted(effectiveStyleIndex);
    let cache: CellCache | undefined;
    let content: CellDocument['content'];
    let requiresExternal = false;

    if (record.formula) {
      let text = record.formula.text;
      if (text === undefined && record.formula.kind === 'shared' && record.formula.si) {
        const master = masters.get(record.formula.si);
        if (!master?.formula?.text) {
          throw new PackageException(`Formula text at ${sheetName}!${record.address} was empty. It was not replaced with zero.`);
        }
        const shifted = sharedFormulaText(master.formula.text, record.row - master.row, record.column - master.column);
        if (shifted === undefined) {
          report.add(exchangeCodes.sharedFormula,
            'Shared formulas that could not be translated keep their master text and are not evaluated.');
          text = master.formula.text;
          requiresExternal = true;
        } else {
          text = shifted;
        }
      }
      if (text === undefined || text.trim() === '') {
        throw new PackageException(`Formula text at ${sheetName}!${record.address} was empty. It was not replaced with zero.`);
      }
      const source = '=' + (text.startsWith('=') ? text.slice(1) : text);
      if (source.length > documentLimits.formulaLength) {
        throw new PackageException(`Formula at ${sheetName}!${record.address} exceeds ${documentLimits.formulaLength} characters. No document was adopted.`);
      }
      content = { kind: 'formula', source };
      requiresExternal = requiresExternal || record.formula.kind === 'array' || formulaNeedsExternalCalculation(source);
      const numeric = record.type === undefined || record.type === 'n' ? record.value : undefined;
      if (numeric !== undefined) {
        let value = parseInvariantNumber(numeric, 'XLSX contains an invalid numeric value/cache.');
        if (use1904 && dateFormatted) value += 1462;
        cache = { kind: 'number', value };
      } else if (record.type === 'b' && record.value !== undefined) {
        cache = { kind: 'boolean', value: record.value === '1' || record.value === 'true' };
      } else if (record.type === 'str' && record.value !== undefined) {
        cache = { kind: 'text', value: record.value };
      } else if (record.type === 's' && record.value !== undefined) {
        cache = { kind: 'text', value: cellTextFromValue(record.value, shared) };
      } else if (record.type === 'e' && record.value !== undefined) {
        cache = { kind: 'error', value: errorValue(record.value) };
      }
    } else {
      const type = record.type;
      if (type === 's') {
        content = { kind: 'text', value: cellTextFromValue(record.value ?? '', shared) };
      } else if (type === 'inlineStr') {
        content = { kind: 'text', value: record.inline ?? '' };
      } else if (type === 'str') {
        content = { kind: 'text', value: record.value ?? '' };
      } else if (type === 'b') {
        content = { kind: 'boolean', value: record.value === '1' || record.value === 'true' };
      } else if (type === 'e') {
        const text = errorValue(record.value ?? '');
        report.add(exchangeCodes.constantError, `A constant error ${text} was imported as text, not as zero.`);
        content = { kind: 'text', value: text };
      } else if (type === 'd') {
        report.add(exchangeCodes.dateSerial, 'Date values are stored as 1900 serial numbers plus their number format.');
        content = { kind: 'number', value: workspaceSerial(record.value ?? '') };
      } else if (record.value !== undefined) {
        let value = parseInvariantNumber(record.value, 'XLSX contains an invalid numeric value/cache.');
        if (use1904 && dateFormatted) value += 1462;
        content = { kind: 'number', value };
      } else {
        content = { kind: 'blank' };
      }
    }

    if (content.kind === 'blank' && style === undefined) continue;
    if (content.kind === 'text' && content.value.length > documentLimits.textLength) {
      throw new PackageException(`Text at ${sheetName}!${record.address} exceeds ${documentLimits.textLength} characters. No document was adopted.`);
    }
    if (record.formula && cache === undefined) {
      report.add(exchangeCodes.missingCache, 'Some imported formulas have no stored cache; no zero is invented.');
    }
    if (requiresExternal) {
      report.add(exchangeCodes.unsupportedFormulas,
        'Unsupported/array/external/name formulas retain source and unverified imported caches; they require Excel for calculation.');
    }
    cells.push({
      id: makeId(), row: record.row, column: record.column, content,
      ...(style ? { style } : {}),
      ...(cache ? { cache } : {}),
      ...(requiresExternal ? { requiresExternalCalculation: true } : {}),
    });
  }

  const merges: CellExtent[] = [];
  for (const merge of childElements(firstChild(raw, 'mergeCells'), 'mergeCell')) {
    const reference = attr(merge, 'ref');
    const area = reference ? parseArea(reference) : undefined;
    if (!area || area.rows * area.columns < 2 || area.row + area.rows - 1 > documentLimits.rows ||
      area.column + area.columns - 1 > documentLimits.columns) {
      throw new PackageException('XLSX contains a merged range outside supported workbook bounds. No document was adopted.');
    }
    merges.push(area);
  }
  const covered = (cell: { row: number; column: number }) => merges.some(area =>
    cell.row >= area.row && cell.row < area.row + area.rows &&
    cell.column >= area.column && cell.column < area.column + area.columns &&
    !(cell.row === area.row && cell.column === area.column));
  const kept = merges.length === 0 ? cells : cells.filter(cell => {
    if (!covered(cell)) return true;
    if (cell.content.kind === 'blank') return true;
    report.add(exchangeCodes.mergedCells, 'Values stored in covered merge cells were discarded; the upper-left value is kept.');
    return false;
  });

  const freeze = readFreeze(raw, report);
  if (firstChild(raw, 'autoFilter')) {
    report.add(exchangeCodes.filterFreeze, 'Excel AutoFilter criteria/sort indicators are omitted; manually hidden rows remain hidden.');
  }
  if (firstChild(raw, 'sheetProtection')) report.add(exchangeCodes.protection, 'Worksheet protection is not retained.');
  if (firstChild(raw, 'dataValidations')) report.add(exchangeCodes.validations, 'Data validations were not imported.');
  if (firstChild(raw, 'conditionalFormatting')) {
    report.add(exchangeCodes.conditionalFormats, 'Conditional formats were not imported.');
  }
  if (firstChild(raw, 'hyperlinks')) report.add(exchangeCodes.hyperlinks, 'Hyperlinks are imported as cell text only; link targets are omitted.');

  return {
    id: makeId(),
    name: sheetName,
    cells: kept,
    rows: dimensions.rows,
    columns: dimensions.columns,
    merges,
    frozenPanes: freeze,
  };
}

function cellTextFromValue(raw: string, shared: readonly string[]): string {
  const index = Number(raw);
  if (!Number.isInteger(index) || index < 0 || index >= shared.length) {
    throw new PackageException('XLSX contains an invalid shared-string reference. No document was adopted.');
  }
  return shared[index]!;
}

function readDimensions(raw: Element, usedMaxRow: number, usedMaxColumn: number,
  styles: ReturnType<typeof readStyles>, report: XlsxReport): {
  rows: SheetDocument['rows'];
  columns: SheetDocument['columns'];
  rowStyleByIndex: ReadonlyMap<number, number>;
  colStyleByIndex: ReadonlyMap<number, number>;
} {
  const rows: { index: number; size?: number; hidden?: boolean; style?: CellStyle }[] = [];
  const rowStyleByIndex = new Map<number, number>();
  const colStyleByIndex = new Map<number, number>();
  const defaultRowHeight = Number(attr(firstChild(raw, 'sheetFormatPr'), 'defaultRowHeight') ?? '15');
  if (Number.isFinite(defaultRowHeight) && Math.abs(defaultRowHeight - 15) > .001) {
    report.add(exchangeCodes.defaultDimensions,
      'Nonstandard sheet default sizes are retained on explicit axes only; unlisted axes use native defaults.');
  }
  for (const row of childElements(firstChild(raw, 'sheetData'), 'row')) {
    const index = Number(attr(row, 'r'));
    if (!Number.isInteger(index) || index < 1 || index > documentLimits.rows) continue;
    const styleText = attr(row, 's');
    const rowStyleIndex = styleText === undefined ? undefined : Number(styleText);
    if (rowStyleIndex !== undefined && (!Number.isInteger(rowStyleIndex) || rowStyleIndex < 0)) {
      throw new PackageException('XLSX style index is invalid. No document was adopted.');
    }
    const style = styles.cellStyle(rowStyleIndex);
    if (style !== undefined && rowStyleIndex !== undefined) rowStyleByIndex.set(index, rowStyleIndex);
    if (style === undefined && index > usedMaxRow) continue;
    const hidden = isTruthyAttribute(row, 'hidden');
    const heightText = attr(row, 'ht');
    let size: number | undefined;
    if (heightText !== undefined) {
      const height = Number(heightText);
      if (Number.isFinite(height) && Math.abs(height - 15) > .001) {
        if (height * 1000 < 4000 || height / .75 > 4096) {
          report.add(exchangeCodes.dimensionLimits, 'Dimensions outside native bounds were clamped.');
        }
        size = Math.min(4096, Math.max(4, height / .75));
      }
    }
    if (size === undefined && !hidden && style === undefined) continue;
    rows.push({ index, ...(size !== undefined ? { size } : {}), hidden, ...(style !== undefined ? { style } : {}) });
  }

  const columns = new Map<number, { index: number; size?: number; hidden?: boolean; style?: CellStyle }>();
  for (const column of childElements(firstChild(raw, 'cols'), 'col')) {
    const min = Number(attr(column, 'min'));
    const declaredMax = Number(attr(column, 'max') ?? attr(column, 'min'));
    if (!Number.isInteger(min) || !Number.isInteger(declaredMax) || min < 1 || declaredMax < min) {
      throw new PackageException('Column dimensions are outside supported workbook limits. No document was adopted.');
    }
    const colStyleText = attr(column, 'style');
    const colStyleIndex = colStyleText === undefined ? undefined : Number(colStyleText);
    if (colStyleIndex !== undefined && (!Number.isInteger(colStyleIndex) || colStyleIndex < 0)) {
      throw new PackageException('XLSX style index is invalid. No document was adopted.');
    }
    const style = styles.cellStyle(colStyleIndex);
    // Whole-sheet ranges such as min=1 max=16384 record a default width. ClosedXML exposes only
    // columns of the used range, so keep the intersection with used/supported columns unless styled.
    const max = Math.min(declaredMax, documentLimits.columns, style !== undefined ? documentLimits.columns : usedMaxColumn);
    if (min > max) continue;
    const hidden = isTruthyAttribute(column, 'hidden');
    const widthText = attr(column, 'width');
    const width = widthText === undefined ? undefined : Number(widthText);
    const pixels = width === undefined || !Number.isFinite(width) ? undefined :
      width < 1 ? width * 12 : width * 7 + 5;
    const differs = width !== undefined && pixels !== undefined && Math.abs(width - 8.43) > .001;
    if (pixels !== undefined && differs && (pixels < 4 || pixels > 4096)) {
      report.add(exchangeCodes.dimensionLimits, 'Dimensions outside native bounds were clamped.');
    }
    const size = differs ? Math.min(4096, Math.max(4, pixels!)) : undefined;
    if (size !== undefined) report.add(exchangeCodes.columnWidth,
      'Column widths use a Calibri 11 / 96-DPI approximation; Excel widths depend on the normal font.');
    for (let index = min; index <= max; index++) {
      if (style !== undefined && colStyleIndex !== undefined) colStyleByIndex.set(index, colStyleIndex);
      if (size === undefined && !hidden && style === undefined) continue;
      columns.set(index, { index, ...(size !== undefined ? { size } : {}), hidden, ...(style !== undefined ? { style } : {}) });
    }
  }
  return { rows, columns: [...columns.values()].sort((left, right) => left.index - right.index), rowStyleByIndex, colStyleByIndex };
}

function readFreeze(raw: Element, report: XlsxReport): SheetDocument['frozenPanes'] {
  const sheetView = firstChild(firstChild(raw, 'sheetViews'), 'sheetView');
  const pane = firstChild(sheetView, 'pane');
  if (!pane) return { rows: 0, columns: 0 };
  const state = attr(pane, 'state');
  if (state !== 'frozen' && state !== 'frozenSplit') {
    report.add(exchangeCodes.freezePanes, 'Independent split windows are omitted.');
    return { rows: 0, columns: 0 };
  }
  const count = (name: string): number => {
    const text = attr(pane, name);
    if (text === undefined) return 0;
    const value = Number(text);
    return Number.isInteger(value) && value >= 0 ? value : -1;
  };
  const rows = count('ySplit');
  const columns = count('xSplit');
  if (rows < 0 || columns < 0 || rows > documentLimits.rows - 1 || columns > documentLimits.columns - 1) {
    report.add(exchangeCodes.freezePanes, 'Frozen pane boundaries outside supported workbook bounds are omitted.');
    return { rows: 0, columns: 0 };
  }
  return { rows, columns };
}

export async function importXlsx(file: Blob, options?: { readonly makeId?: IdFactory }): Promise<XlsxImport> {
  const makeId = options?.makeId ?? (() => crypto.randomUUID());
  const report = new XlsxReport();
  const { parts, entries } = await readPackage(file);
  scanEntries(entries, report);
  checkContentType(parts);
  const workbookBytes = parts.get('xl/workbook.xml');
  const relationshipBytes = parts.get('xl/_rels/workbook.xml.rels');
  if (!workbookBytes) {
    throw new PackageException('This file is not an XLSX workbook. No document was adopted.');
  }
  if (!relationshipBytes) throw new PackageException('XLSX relationships are missing. No document was adopted.');

  const workbook = parseXmlPart(workbookBytes, 'The XLSX workbook part');
  const relationships = parseXmlPart(relationshipBytes, 'The XLSX relationships part');
  const targets = new Map<string, string>();
  for (const relationship of childElements(relationships.root, 'Relationship')) {
    const id = attr(relationship, 'Id');
    const target = attr(relationship, 'Target');
    if (!id || !target) continue;
    if (attr(relationship, 'TargetMode') === 'External') continue;
    targets.set(id, resolveTarget(target));
  }

  const use1904 = isTruthyAttribute(firstChild(workbook.root, 'workbookPr') ?? workbook.root, 'date1904');
  if (use1904) {
    report.add(exchangeCodes.dateSystem,
      '1904 date-formatted numeric cells/caches are converted to 1900 serials (+1462 days). Numeric constants in formula source are retained; review date arithmetic before recalculating. Time-only values are unchanged.');
  }

  const sheetElements = childElements(firstChild(workbook.root, 'sheets'), 'sheet');
  if (sheetElements.length === 0) {
    throw new PackageException('XLSX has no worksheets. No document was adopted.');
  }
  if (sheetElements.length > documentLimits.sheets) {
    throw new PackageException(`The workbook has more than ${documentLimits.sheets} sheets. No document was adopted.`);
  }
  const names = childElements(firstChild(workbook.root, 'definedNames'), 'definedName')
    .map(name => textOf(name) ?? '').filter(name => name.length > 0);
  if (names.length > 0) {
    report.add(exchangeCodes.definedNames, 'Defined names were not imported: ' +
      names.slice(0, maxReportedNames).join(', ') + (names.length > maxReportedNames ? ` and ${names.length - maxReportedNames} more` : '') + '.');
  }

  const shared = readSharedStrings(parts.get('xl/sharedstrings.xml'), report);
  const styles = readStyles(parts.get('xl/styles.xml'), report);
  const counters = { cells: 0, formulas: 0 };
  const sheets: SheetDocument[] = [];
  for (const sheet of sheetElements) {
    const name = attr(sheet, 'name');
    const id = relationshipId(sheet);
    if (!name) throw new PackageException('XLSX sheet name is missing. No document was adopted.');
    const target = id ? targets.get(id) : undefined;
    if (!target) throw new PackageException(`XLSX worksheet relationship for ${name} is missing. No document was adopted.`);
    const bytes = parts.get(target.toLowerCase());
    if (!bytes) throw new PackageException(`Worksheet part ${target} is missing. No document was adopted.`);
    if (attr(sheet, 'state') === 'hidden' || attr(sheet, 'state') === 'veryHidden') {
      report.add(exchangeCodes.hiddenSheet, `${name} is hidden. The import keeps its cells and does not store visibility.`);
    }
    const parsed = parseXmlPart(bytes, `Worksheet ${name}`);
    sheets.push(readWorksheet(parsed.root, name, shared, styles, use1904, report, makeId, counters));
  }

  if (counters.formulas > 0) {
    report.add(exchangeCodes.unevaluatedFormulas,
      'Formula text and caches were imported. Formulas were not evaluated. Unsupported formula text is not replaced with zero.');
  }
  report.add(exchangeCodes.newIds, 'XLSX has no stable sheet or cell ids. New ids were assigned and are kept only in .exlt.');

  const document: WorkbookDocument = {
    id: makeId(),
    formatVersion: 1,
    dateSystem: '1900',
    sheets,
    names: [],
  };
  try {
    validateWorkbook(document);
  } catch (error) {
    if (error instanceof DocumentError) {
      throw new PackageException(error.message + ' No document was adopted.');
    }
    throw error;
  }
  return { document, report };
}
