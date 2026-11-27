import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { findCell, type SheetDocument } from '../../src/core/workbook';
import { loadNativePackage, saveNativePackage } from '../../src/persistence/native';
import { writeProbeParts } from '../../src/persistence/bounded-zip';
import { importXlsx } from '../../src/persistence/xlsx';

async function fixture(path: string): Promise<Blob> {
  return new Blob([await readFile(new URL(path, import.meta.url))]);
}

const contentTypes = '<?xml version="1.0" encoding="UTF-8"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
  '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
  '</Types>';
const relationships = '<?xml version="1.0" encoding="UTF-8"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
  '</Relationships>';

function workbookXml(sheets: string, extra = '', properties = ''): string {
  return '<?xml version="1.0" encoding="UTF-8"?>' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    `<workbookPr${properties}/><sheets>${sheets}</sheets>${extra}</workbook>`;
}

function worksheetXml(body: string, attributes = ''): string {
  return '<?xml version="1.0" encoding="UTF-8"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    (attributes ? `<sheetViews><sheetView workbookViewId="0">${attributes}</sheetView></sheetViews>` : '') +
    `<sheetData>${body}</sheetData></worksheet>`;
}

async function packageOf(files: Record<string, string>): Promise<Blob> {
  return writeProbeParts(new Map(Object.entries(files).map(([name, text]) =>
    [name, new TextEncoder().encode(text)])));
}

function simplePackage(sheet: string, workbookExtra = '', properties = ''): Promise<Blob> {
  return packageOf({
    '[Content_Types].xml': contentTypes,
    'xl/workbook.xml': workbookXml('<sheet name="Sheet1" sheetId="1" r:id="rId1"/>', workbookExtra, properties),
    'xl/_rels/workbook.xml.rels': relationships,
    'xl/worksheets/sheet1.xml': sheet,
  });
}

function cell(sheet: SheetDocument, row: number, column: number) {
  return findCell(sheet, { row, column });
}

it('imports the desktop extended style fixture with gradients, borders, rotation and freeze', async () => {
  const { document, report } = await importXlsx(await fixture('../fixtures/interchange/DesktopStyles.xlsx'));
  expect(document.sheets.map(sheet => sheet.name)).toEqual(['Styles']);
  const sheet = document.sheets[0]!;
  expect(sheet.frozenPanes).toEqual({ rows: 1, columns: 0 });
  expect(sheet.columns).toEqual([{ index: 1, size: 68.984375, hidden: false }]);

  expect(cell(sheet, 1, 1)?.content).toEqual({ kind: 'number', value: 60 });
  expect(cell(sheet, 1, 1)?.style).toEqual({
    numberFormat: 'yyyy-mm-dd',
    fillRgb: 'FFFFFF',
    alignment: 'distributed',
    fillDetails: { pattern: 'none', gradientRgb: '2288AA', shading: 'diagonalDown', variant: 3 },
    borderDetails: {
      top: { style: 'slantDashDot', rgb: '112233' },
      right: { style: 'double' },
      bottom: { style: 'mediumDashDotDot', rgb: '445566' },
      diagonalUp: { style: 'dotted', rgb: '778899' },
      diagonalDown: { style: 'dotted', rgb: '778899' },
    },
    textRotation: -45,
    indent: 3,
    readingOrder: 'rtl',
    justifyLastLine: true,
  });
  const stacked = cell(sheet, 2, 1);
  expect(stacked?.content).toEqual({ kind: 'text', value: 'stacked' });
  expect(stacked?.style?.textRotation).toBe(255);
  expect(stacked?.style?.fillDetails).toEqual({ pattern: 'none', gradientRgb: '2288AA', shading: 'center', variant: 2 });

  expect(report.issues.map(issue => issue.code)).toEqual(['axis-styles', 'column-width', 'new-ids']);
  expect(report.has('merged-cells')).toBe(false);
});

it('converts 1904 date-formatted numbers and caches by 1462 while keeping clocks, elapsed time and plain numbers', async () => {
  const { document, report } = await importXlsx(await fixture('../fixtures/formats/Desktop1904.xlsx'));
  const sheet = document.sheets[0]!;
  expect([1, 2, 3, 4, 5].map(row => cell(sheet, row, 1)?.content)).toEqual([
    { kind: 'number', value: 1463.5 },
    { kind: 'number', value: 1.5 },
    { kind: 'number', value: 1.5 },
    { kind: 'number', value: 1.5 },
    { kind: 'number', value: 1463.5 },
  ]);
  expect(report.has('date-system')).toBe(true);

  const styles = '<?xml version="1.0" encoding="UTF-8"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>' +
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
    '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>' +
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>';
  const sheetXml = worksheetXml('<row r="1"><c r="A1" s="1"><f>DATE(2020,1,1)</f><v>1.5</v></c></row>');
  const { document: synthetic } = await importXlsx(await packageOf({
    '[Content_Types].xml': contentTypes,
    'xl/workbook.xml': workbookXml('<sheet name="Sheet1" sheetId="1" r:id="rId1"/>', '', ' date1904="1"'),
    'xl/_rels/workbook.xml.rels': relationships,
    'xl/styles.xml': styles,
    'xl/worksheets/sheet1.xml': sheetXml,
  }));
  expect(cell(synthetic.sheets[0]!, 1, 1)?.content).toEqual({ kind: 'formula', source: '=DATE(2020,1,1)' });
  expect(cell(synthetic.sheets[0]!, 1, 1)?.cache).toEqual({ kind: 'number', value: 1463.5 });
});

it('imports scalar types, explicit empty text, constant errors and typed formula caches', async () => {
  const shared = '<?xml version="1.0" encoding="UTF-8"?>' +
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>Hello</t></si></sst>';
  const sheet = worksheetXml('<row r="1">' +
    '<c r="A1"><v>42</v></c>' +
    '<c r="B1" t="s"><v>0</v></c>' +
    '<c r="C1" t="b"><v>1</v></c>' +
    '<c r="D1" t="e"><v>#DIV/0!</v></c>' +
    '<c r="E1" t="inlineStr"><is><t></t></is></c>' +
    '<c r="F1"><f>SUM(A1:A1)</f><v>42</v></c>' +
    '<c r="G1"><f>UNKNOWN(1)</f></c>' +
    '<c r="H1" t="str"><f>"x"</f><v>x</v></c>' +
    '<c r="I1" t="b"><f>IF(A1=42,TRUE,FALSE)</f><v>1</v></c>' +
    '<c r="J1" t="e"><f>1/0</f><v>#DIV/0!</v></c>' +
    '</row>');
  const { document, report } = await importXlsx(await packageOf({
    '[Content_Types].xml': contentTypes,
    'xl/workbook.xml': workbookXml('<sheet name="Sheet1" sheetId="1" r:id="rId1"/>'),
    'xl/_rels/workbook.xml.rels': relationships,
    'xl/sharedStrings.xml': shared,
    'xl/worksheets/sheet1.xml': sheet,
  }));
  const target = document.sheets[0]!;
  expect(cell(target, 1, 1)?.content).toEqual({ kind: 'number', value: 42 });
  expect(cell(target, 1, 2)?.content).toEqual({ kind: 'text', value: 'Hello' });
  expect(cell(target, 1, 3)?.content).toEqual({ kind: 'boolean', value: true });
  expect(cell(target, 1, 4)?.content).toEqual({ kind: 'text', value: '#DIV/0!' });
  expect(cell(target, 1, 5)?.content).toEqual({ kind: 'text', value: '' });
  expect(cell(target, 1, 6)?.content).toEqual({ kind: 'formula', source: '=SUM(A1:A1)' });
  expect(cell(target, 1, 6)?.cache).toEqual({ kind: 'number', value: 42 });
  expect(cell(target, 1, 6)?.requiresExternalCalculation).toBeUndefined();
  expect(cell(target, 1, 7)?.content).toEqual({ kind: 'formula', source: '=UNKNOWN(1)' });
  expect(cell(target, 1, 7)?.cache).toBeUndefined();
  expect(cell(target, 1, 7)?.requiresExternalCalculation).toBe(true);
  expect(cell(target, 1, 8)?.cache).toEqual({ kind: 'text', value: 'x' });
  expect(cell(target, 1, 9)?.cache).toEqual({ kind: 'boolean', value: true });
  expect(cell(target, 1, 10)?.cache).toEqual({ kind: 'error', value: '#DIV/0!' });

  expect(report.has('constant-error')).toBe(true);
  expect(report.has('unsupported-formulas')).toBe(true);
  expect(report.has('missing-cache')).toBe(true);
  expect(report.has('unevaluated-formulas')).toBe(true);
  expect(report.has('new-ids')).toBe(true);
});

it('resolves shared formulas by shifting relative references for each dependent cell', async () => {
  const sheet = worksheetXml('<row r="1">' +
    '<c r="A1"><f t="shared" ref="A1:A3" si="0">B1+1</f><v>2</v></c>' +
    '<c r="A2"><f t="shared" si="0"/><v>3</v></c>' +
    '<c r="A3"><f t="shared" si="0"/><v>4</v></c>' +
    '</row>');
  const { document } = await importXlsx(await simplePackage(sheet));
  const target = document.sheets[0]!;
  expect(cell(target, 1, 1)?.content).toEqual({ kind: 'formula', source: '=B1+1' });
  expect(cell(target, 2, 1)?.content).toEqual({ kind: 'formula', source: '=B2+1' });
  expect(cell(target, 3, 1)?.content).toEqual({ kind: 'formula', source: '=B3+1' });
});

it('keeps the merged anchor and reports covered values it discards', async () => {
  const merged = worksheetXml('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>' +
    '<row r="2"><c r="B2"/></row>').replace('</worksheet>',
    '<mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells></worksheet>');
  const { document, report } = await importXlsx(await simplePackage(merged));
  const target = document.sheets[0]!;
  expect(target.merges).toEqual([{ row: 1, column: 1, rows: 2, columns: 2 }]);
  expect(target.cells.filter(entry => entry.row === 1).map(entry => entry.column)).toEqual([1]);
  expect(cell(target, 1, 1)?.content).toEqual({ kind: 'number', value: 1 });
  expect(report.has('merged-cells')).toBe(true);
});

it('reports every omitted workbook object family it can detect', async () => {
  const sheet = worksheetXml('<row r="1"><c r="A1"><v>1</v></c></row>',
    '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="split"/>').replace('</worksheet>',
    '<autoFilter ref="A1:B2"/><sheetProtection sheet="1"/>' +
    '<dataValidations count="1"><dataValidation sqref="A1"><whole/></dataValidation></dataValidations>' +
    '<conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1"/></conditionalFormatting>' +
    '<hyperlinks><hyperlink ref="A1"/></hyperlinks></worksheet>');
  const workbookExtra = '<definedNames><definedName name="Total">Sheet1!$A$1</definedName></definedNames>';
  const { document, report } = await importXlsx(await packageOf({
    '[Content_Types].xml': contentTypes,
    'xl/workbook.xml': workbookXml(
      '<sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Hidden" sheetId="2" state="hidden" r:id="rId2"/>',
      workbookExtra),
    'xl/_rels/workbook.xml.rels': '<?xml version="1.0" encoding="UTF-8"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="worksheet" Target="worksheets/sheet2.xml"/></Relationships>',
    'xl/worksheets/sheet1.xml': sheet,
    'xl/worksheets/sheet2.xml': worksheetXml('<row r="1"><c r="A1"><v>5</v></c></row>'),
    'xl/tables/table1.xml': '<table/>',
    'xl/charts/chart1.xml': '<chartSpace/>',
    'xl/comments1.xml': '<comments/>',
    'xl/drawings/drawing1.xml': '<wsDr/>',
    'xl/pivotTables/pivotTable1.xml': '<pivotTableDefinition/>',
    'xl/externalLinks/externalLink1.xml': '<externalLink/>',
  }));
  expect(document.sheets.map(entry => entry.name)).toEqual(['Sheet1', 'Hidden']);
  expect(document.names).toEqual([]);
  expect(cell(document.sheets[1]!, 1, 1)?.content).toEqual({ kind: 'number', value: 5 });
  for (const code of ['defined-names', 'hidden-sheet', 'filter-freeze', 'validations',
    'conditional-formats', 'hyperlinks', 'protection', 'tables', 'charts', 'comments',
    'drawings', 'pivots', 'external-links', 'freeze-panes']) {
    expect(report.has(code), code).toBe(true);
  }
});

it('saves every imported fixture as an equal native package for browser storage', async () => {
  for (const path of ['../fixtures/interchange/DesktopStyles.xlsx', '../fixtures/formats/Desktop1904.xlsx',
    '../fixtures/sales/showcase_data.xlsx']) {
    const { document } = await importXlsx(await fixture(path));
    const saved = await saveNativePackage(document);
    expect(await loadNativePackage(saved), path).toEqual(document);
  }
});

it('imports a cell at XFD1048576 and keeps a whole-sheet default column width empty', async () => {
  const sheet = worksheetXml('<row r="1048576"><c r="XFD1048576"><v>42</v></c></row>').replace(
    '<sheetData>',
    '<cols><col min="1" max="16384" width="8.43"/></cols><sheetData>');
  const { document, report } = await importXlsx(await simplePackage(sheet));
  expect(cell(document.sheets[0]!, 1_048_576, 16_384)?.content).toEqual({ kind: 'number', value: 42 });
  expect(document.sheets[0]!.columns).toEqual([]);
  expect(report.has('column-width')).toBe(false);
});

it('rejects encrypted, macro-enabled, out-of-bounds and malformed packages before adoption', async () => {
  const encrypted = new Blob([new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0])]);
  await expect(importXlsx(encrypted)).rejects.toThrow(/Encrypted\/password-protected or legacy Excel files/);

  await expect(importXlsx(await packageOf({
    '[Content_Types].xml': contentTypes,
    'xl/workbook.xml': workbookXml('<sheet name="Sheet1" sheetId="1" r:id="rId1"/>'),
    'xl/_rels/workbook.xml.rels': relationships,
    'xl/worksheets/sheet1.xml': worksheetXml(''),
    'xl/vbaProject.bin': 'macro',
  }))).rejects.toThrow(/Macro-enabled Excel files are unsupported/);

  await expect(importXlsx(await simplePackage(worksheetXml('<row r="1048577"><c r="A1048577"><v>1</v></c></row>'))))
    .rejects.toThrow(/No document was adopted/);

  await expect(importXlsx(await simplePackage(worksheetXml('<row r="1"><c r="XFE1"><v>1</v></c></row>'))))
    .rejects.toThrow(/No document was adopted/);

  await expect(importXlsx(new Blob([new TextEncoder().encode('not a zip')])))
    .rejects.toThrow(/could not be read as a ZIP package/);

  await expect(importXlsx(await simplePackage(worksheetXml('<row r="1"><c r="A1"><f/></c></row>'))))
    .rejects.toThrow(/Formula text at Sheet1!A1 was empty/);

  await expect(importXlsx(await simplePackage(worksheetXml('').replace('<worksheet',
    '<!DOCTYPE worksheet [<!ENTITY inflated "repeated">]><worksheet'))))
    .rejects.toThrow(/document type declaration/);

  await expect(importXlsx(await simplePackage('<worksheet><sheetData>')))
    .rejects.toThrow(/not valid XML/);

  const mergedOutside = worksheetXml('<row r="1"><c r="A1"><v>1</v></c></row>').replace('</worksheet>',
    '<mergeCells count="1"><mergeCell ref="A1:B1048577"/></mergeCells></worksheet>');
  await expect(importXlsx(await simplePackage(mergedOutside)))
    .rejects.toThrow(/merged range outside supported workbook bounds/);
});
