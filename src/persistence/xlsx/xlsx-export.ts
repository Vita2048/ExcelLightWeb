import { BlobWriter, TextReader, ZipWriter } from '@zip.js/zip.js';
import type { CellCache, CellDocument, CellStyle, SheetDocument, WorkbookDocument } from '../../core/workbook';
import { columnLetters } from '../../core/formula-references';
import { validateWorkbook } from '../../core/workbook-validation';
import { PackageException } from '../native';
import { createStyleWriter } from './xlsx-styles';
import { escapeAttribute, escapeText } from './xlsx-xml';
import { exchangeCodes, XlsxReport } from './xlsx-report';
import { addXlfnPrefixes } from './xlsx-formulas';

const spreadsheetNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const relationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const packageRelationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships';
const workbookContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
const worksheetContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const stylesContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
const sharedStringsContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';
const exportMimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const maxRowPoints = 409;
const maxColumnWidth = 255;

export interface XlsxExport {
  readonly blob: Blob;
  readonly report: XlsxReport;
}

function address(row: number, column: number): string {
  return columnLetters(column) + row;
}

function join(entries: readonly string[]): string {
  return entries.slice(0, 12).join(', ') + (entries.length > 12 ? ` and ${entries.length - 12} more` : '');
}

/** Desktop XlsxExchange.Inspect: reviewable losses before writing an export. */
export function inspectXlsxExport(document: WorkbookDocument): XlsxReport {
  validateWorkbook(document);
  const report = new XlsxReport();
  report.add(exchangeCodes.stableIds, 'Stable sheet and cell ids are not written to XLSX.');
  report.add(exchangeCodes.unevaluatedFormulas,
    'Formula text and stored caches are written. Formulas are not evaluated. Unsupported formula text is not replaced with zero.');
  report.add(exchangeCodes.omittedObjects,
    'Charts, pivots, comments, tables, defined names, validations, conditional formats, drawings and macros are not exported.');
  if (document.names.some(name => !name.isTable)) {
    report.add(exchangeCodes.definedNames, 'Workbook named ranges and comments are omitted from XLSX. Native .exlt retains them.');
  }
  if (document.names.some(name => name.isTable)) {
    report.add(exchangeCodes.tables,
      'Lightweight table names, explicit extents and header contracts are omitted from XLSX. Cell values remain; native .exlt retains definitions.');
  }
  if ((document.pivots?.length ?? 0) > 0) {
    report.add(exchangeCodes.pivots,
      'PivotTable definitions (pivotTableDefinition/pivotCache) are omitted from XLSX; materialized header/value cells and number formats are exported. Native .exlt retains live PivotTables.');
  }
  if (document.sheets.some(sheet => sheet.cells.some(cell => cell.requiresExternalCalculation))) {
    report.add(exchangeCodes.unsupportedFormulas,
      'Unsupported formula sources/caches are retained without evaluation. Array/spill semantics are not exported; formula text is written as a scalar formula. Native .exlt retains external-calculation metadata.');
  }
  if (document.sheets.some(sheet => sheet.filter !== undefined)) {
    report.add(exchangeCodes.filterFreeze,
      'Selected-value AutoFilter criteria and sort indicators are native-only and omitted from XLSX; manual hiding and frozen panes are exported.');
  }
  report.add(exchangeCodes.pythonSnapshots,
    'Only materialized scalar/table snapshot cells are exported. Python source, bindings, trust and figures are native-only. Export never executes Python.');
  if (document.sheets.some(sheet => sheet.rows.some(row => (row.size ?? 0) > maxRowPoints / .75) ||
    sheet.columns.some(column => (column.size ?? 0) > maxColumnWidth * 7 + 5))) {
    report.add(exchangeCodes.dimensionLimits,
      'Row heights above 409 points and column widths above 255 characters are clamped for Excel. Native .exlt retains the original dimensions.');
  }
  if (document.sheets.some(sheet => sheet.columns.some(column => column.size !== undefined))) {
    report.add(exchangeCodes.columnWidth,
      'Column widths use a Calibri 11 / 96-DPI approximation; Excel widths depend on the normal font.');
  }
  report.add(exchangeCodes.cacheFreshness,
    'Stored formula caches may be stale or unverified. Export does not recalculate; Excel is asked to recalculate on open.');
  const missing: string[] = [];
  for (const sheet of document.sheets) {
    for (const cell of sheet.cells) {
      if (cell.content.kind === 'formula' && cell.cache === undefined) {
        missing.push(sheet.name + '!' + address(cell.row, cell.column));
      }
    }
  }
  if (missing.length > 0) {
    report.add(exchangeCodes.missingCache,
      'No cache is stored for ' + join(missing) + '. Export will not invent 0.');
  }
  return report;
}

class SharedStrings {
  readonly #indexes = new Map<string, number>();
  readonly items: string[] = [];

  index(value: string): number {
    const existing = this.#indexes.get(value);
    if (existing !== undefined) return existing;
    const index = this.items.length;
    this.items.push(value);
    this.#indexes.set(value, index);
    return index;
  }

  get xml(): string {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<sst xmlns="${spreadsheetNamespace}" count="${this.items.length}" uniqueCount="${this.items.length}">` +
      this.items.map(value => `<si><t${/^\s|\s$/u.test(value) ? ' xml:space="preserve"' : ''}>` +
        escapeText(value) + '</t></si>').join('') + '</sst>';
  }
}

function cacheValue(cache: CellCache): { readonly value: string; readonly type?: string } {
  switch (cache.kind) {
    case 'number':
      return { value: cache.serializedText !== undefined && Number(cache.serializedText) === cache.value
        ? cache.serializedText : String(cache.value) };
    case 'boolean':
      return { value: cache.value ? '1' : '0', type: 'b' };
    case 'text':
      return { value: escapeText(cache.value), type: 'str' };
    case 'error':
      return { value: escapeText(cache.value), type: 'e' };
  }
}

function cellXml(cell: CellDocument, styleIndex: number, strings: SharedStrings, report: XlsxReport): string {
  const reference = address(cell.row, cell.column);
  const style = cell.style !== undefined ? ` s="${styleIndex}"` : '';
  const content = cell.content;
  if (content.kind === 'formula') {
    const rawFormula = content.source.startsWith('=') ? content.source.slice(1) : content.source;
    const formula = `<f>${escapeText(addXlfnPrefixes(rawFormula))}</f>`;
    if (!cell.cache) return `<c r="${reference}"${style}>${formula}</c>`;
    const cache = cacheValue(cell.cache);
    return `<c r="${reference}"${style}${cache.type ? ` t="${cache.type}"` : ''}>${formula}` +
      `<v>${cache.value}</v></c>`;
  }
  if (content.kind === 'python') {
    report.add(exchangeCodes.pythonSnapshots,
      'Only materialized scalar/table snapshot cells are exported. Python source, bindings, trust and figures are native-only. Export never executes Python.');
    if (!cell.cache) return `<c r="${reference}"${style}/>`;
    const cache = cell.cache;
    if (cache.kind === 'text') return `<c r="${reference}"${style} t="s"><v>${strings.index(cache.value)}</v></c>`;
    const value = cacheValue(cache);
    return `<c r="${reference}"${style}${value.type ? ` t="${value.type}"` : ''}><v>${value.value}</v></c>`;
  }
  switch (content.kind) {
    case 'blank':
      if (cell.cache?.kind === 'error') {
        return `<c r="${reference}"${style} t="e"><v>${escapeText(cell.cache.value)}</v></c>`;
      }
      return `<c r="${reference}"${style}/>`;
    case 'number':
      return `<c r="${reference}"${style}><v>${String(content.value)}</v></c>`;
    case 'boolean':
      return `<c r="${reference}"${style} t="b"><v>${content.value ? '1' : '0'}</v></c>`;
    case 'text':
      return `<c r="${reference}"${style} t="s"><v>${strings.index(content.value)}</v></c>`;
  }
}

function paneXml(sheet: SheetDocument): string {
  const { rows, columns } = sheet.frozenPanes;
  if (rows === 0 && columns === 0) return '<sheetView workbookViewId="0"/>';
  const activePane = rows > 0 && columns > 0 ? 'bottomRight' : rows > 0 ? 'bottomLeft' : 'topRight';
  return `<sheetView workbookViewId="0"><pane${columns > 0 ? ` xSplit="${columns}"` : ''}` +
    `${rows > 0 ? ` ySplit="${rows}"` : ''} topLeftCell="${address(rows + 1, columns + 1)}"` +
    ` activePane="${activePane}" state="frozen"/></sheetView>`;
}

function worksheetXml(sheet: SheetDocument, styleIndexes: (style: CellStyle | undefined) => number,
  strings: SharedStrings, report: XlsxReport): string {
  const dimensionRows = new Map(sheet.rows.map(row => [row.index, row]));
  const dimensionColumns = new Map(sheet.columns.map(column => [column.index, column]));
  const cellsByRow = new Map<number, CellDocument[]>();
  let maxRow = 1;
  let maxColumn = 1;
  for (const cell of sheet.cells) {
    const group = cellsByRow.get(cell.row);
    if (group) group.push(cell); else cellsByRow.set(cell.row, [cell]);
    if (cell.row > maxRow) maxRow = cell.row;
    if (cell.column > maxColumn) maxColumn = cell.column;
  }
  const rowIndexes = new Set<number>([...cellsByRow.keys(), ...dimensionRows.keys()]);
  const rowsXml = [...rowIndexes].sort((left, right) => left - right).map(index => {
    const dimension = dimensionRows.get(index);
    const attributes: string[] = [`r="${index}"`];
    if (dimension?.size !== undefined) {
      attributes.push(`ht="${Math.min(maxRowPoints, dimension.size * .75)}"`, 'customHeight="1"');
    }
    if (dimension?.hidden) attributes.push('hidden="1"');
    if (dimension?.style !== undefined) {
      attributes.push(`s="${styleIndexes(dimension.style)}"`, 'customFormat="1"');
    }
    const rowCells = (cellsByRow.get(index) ?? [])
      .sort((left, right) => left.column - right.column)
      .map(cell => cellXml(cell, styleIndexes(cell.style), strings, report)).join('');
    return `<row ${attributes.join(' ')}>${rowCells}</row>`;
  }).join('');

  const colGroups: { min: number; max: number; width?: number; hidden: boolean; styleIndex?: number }[] = [];
  for (const column of [...dimensionColumns.values()].sort((left, right) => left.index - right.index)) {
    const width = column.size === undefined ? undefined :
      Math.min(maxColumnWidth, column.size < 12 ? column.size / 12 : (column.size - 5) / 7);
    const hidden = !!column.hidden;
    const styleIndex = column.style !== undefined ? styleIndexes(column.style) : undefined;
    const prev = colGroups.at(-1);
    if (prev && column.index === prev.max + 1 && prev.width === width && prev.hidden === hidden && prev.styleIndex === styleIndex) {
      prev.max = column.index;
    } else {
      colGroups.push({
        min: column.index,
        max: column.index,
        ...(width !== undefined ? { width } : {}),
        hidden,
        ...(styleIndex !== undefined ? { styleIndex } : {}),
      });
    }
  }
  const columnsXml = colGroups.map(group =>
    `<col min="${group.min}" max="${group.max}"` +
    (group.width !== undefined ? ` width="${group.width}" customWidth="1"` : '') +
    (group.hidden ? ' hidden="1"' : '') +
    (group.styleIndex !== undefined ? ` style="${group.styleIndex}"` : '') + '/>',
  ).join('');

  for (const row of sheet.rows) if (row.index > maxRow) maxRow = row.index;
  for (const column of sheet.columns) if (column.index > maxColumn) maxColumn = column.index;
  const merges = sheet.merges.map(merge => `<mergeCell ref="${address(merge.row, merge.column)}:` +
    `${address(merge.row + merge.rows - 1, merge.column + merge.columns - 1)}"/>`).join('');

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<worksheet xmlns="${spreadsheetNamespace}" xmlns:r="${relationshipNamespace}">` +
    `<dimension ref="A1:${address(maxRow, maxColumn)}"/>` +
    `<sheetViews>${paneXml(sheet)}</sheetViews>` +
    '<sheetFormatPr defaultRowHeight="15"/>' +
    (columnsXml ? `<cols>${columnsXml}</cols>` : '') +
    `<sheetData>${rowsXml}</sheetData>` +
    (merges ? `<mergeCells count="${sheet.merges.length}">${merges}</mergeCells>` : '') +
    '</worksheet>';
}

function contentTypesXml(sheetCount: number, hasSharedStrings: boolean): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    `<Override PartName="/xl/workbook.xml" ContentType="${workbookContentType}"/>` +
    `<Override PartName="/xl/styles.xml" ContentType="${stylesContentType}"/>` +
    (hasSharedStrings ? `<Override PartName="/xl/sharedStrings.xml" ContentType="${sharedStringsContentType}"/>` : '') +
    Array.from({ length: sheetCount }, (_, index) =>
      `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="${worksheetContentType}"/>`).join('') +
    '</Types>';
}

function workbookXml(document: WorkbookDocument): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<workbook xmlns="${spreadsheetNamespace}" xmlns:r="${relationshipNamespace}">` +
    '<workbookPr date1904="0"/><bookViews><workbookView activeTab="0"/></bookViews><sheets>' +
    document.sheets.map((sheet, index) =>
      `<sheet name="${escapeAttribute(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join('') +
    '</sheets>' +
    '<calcPr calcId="0" calcMode="manual" fullCalcOnLoad="1"/>' +
    '</workbook>';
}

function relationships(count: number, hasSharedStrings: boolean): string {
  let id = count + 1;
  const extra: string[] = [
    `<Relationship Id="rId${id++}" Type="${relationshipNamespace}/styles" Target="styles.xml"/>`,
  ];
  if (hasSharedStrings) {
    extra.push(`<Relationship Id="rId${id++}" Type="${relationshipNamespace}/sharedStrings" Target="sharedStrings.xml"/>`);
  }
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<Relationships xmlns="${packageRelationshipNamespace}">` +
    Array.from({ length: count }, (_, index) =>
      `<Relationship Id="rId${index + 1}" Type="${relationshipNamespace}/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).join('') +
    extra.join('') + '</Relationships>';
}

export async function writeXlsx(document: WorkbookDocument): Promise<XlsxExport> {
  validateWorkbook(document);
  const report = inspectXlsxExport(document);
  const styles = createStyleWriter();
  const strings = new SharedStrings();
  const sheets = document.sheets.map(sheet => worksheetXml(sheet, style => styles.indexOf(style), strings, report));
  const hasSharedStrings = strings.items.length > 0;

  const writer = new ZipWriter(new BlobWriter(exportMimeType), { useWebWorkers: false, zip64: false });
  try {
    await writer.add('[Content_Types].xml', new TextReader(contentTypesXml(document.sheets.length, hasSharedStrings)));
    await writer.add('_rels/.rels', new TextReader('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<Relationships xmlns="${packageRelationshipNamespace}">` +
      `<Relationship Id="rId1" Type="${relationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>` +
      '</Relationships>'));
    await writer.add('xl/workbook.xml', new TextReader(workbookXml(document)));
    await writer.add('xl/_rels/workbook.xml.rels', new TextReader(relationships(document.sheets.length, hasSharedStrings)));
    await writer.add('xl/styles.xml', new TextReader(styles.toXml()));
    if (hasSharedStrings) await writer.add('xl/sharedStrings.xml', new TextReader(strings.xml));
    for (const [index, xml] of sheets.entries()) {
      await writer.add(`xl/worksheets/sheet${index + 1}.xml`, new TextReader(xml));
    }
  } catch (error) {
    await writer.close().catch(() => {});
    if (error instanceof PackageException) throw error;
    throw error;
  }
  const blob = await writer.close();
  return { blob, report };
}
