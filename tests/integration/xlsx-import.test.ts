import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { runNativeGeneration } from '../../src/calculation/native-generation';
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

  expect(report.issues.map(issue => issue.code)).toEqual(['column-width', 'new-ids']);
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

it('imports Excel Tables and worksheet/table-backed PivotTables cleanly and calculates structured references natively', async () => {
  const pkg = await packageOf({
    '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '</Types>',
    'xl/styles.xml': '<?xml version="1.0" encoding="UTF-8"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="$#,##0"/></numFmts>' +
      '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>' +
      '</styleSheet>',
    'xl/workbook.xml': workbookXml(
      '<sheet name="Dashboard" sheetId="1" r:id="rId1"/>' +
      '<sheet name="Orders" sheetId="2" r:id="rId2"/>' +
      '<sheet name="Pivot" sheetId="3" r:id="rId3"/>',
    ),
    'xl/_rels/workbook.xml.rels': '<?xml version="1.0" encoding="UTF-8"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>' +
      '</Relationships>',
    'xl/worksheets/sheet1.xml': worksheetXml(
      '<row r="9"><c r="A9" t="inlineStr"><is><t>North</t></is></c><c r="B9"><f>SUMIFS(Orders[Revenue],Orders[Region],A9)</f><v>100</v></c></row>' +
      '<row r="10"><c r="A10" t="inlineStr"><is><t>South</t></is></c><c r="B10"><f>SUMIFS(Orders[Revenue],Orders[Region],A10)</f><v>200</v></c></row>' +
      '<row r="13"><c r="A13" t="inlineStr"><is><t>Total</t></is></c><c r="B13"><f>SUM(B9:B10)</f><v>300</v></c></row>',
    ),
    'xl/worksheets/sheet2.xml': worksheetXml(
      '<row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c><c r="B1" t="inlineStr"><is><t>Category</t></is></c><c r="C1" t="inlineStr"><is><t>Units</t></is></c><c r="D1" t="inlineStr"><is><t>Unit Price</t></is></c><c r="E1" t="inlineStr"><is><t>Revenue</t></is></c></row>' +
      '<row r="2"><c r="A2" t="inlineStr"><is><t>North</t></is></c><c r="B2" t="inlineStr"><is><t>Furniture</t></is></c><c r="C2"><v>2</v></c><c r="D2"><v>50</v></c><c r="E2"><f>Orders[[#This Row],[Units]]*Orders[[#This Row],[Unit Price]]</f><v>100</v></c></row>' +
      '<row r="3"><c r="A3" t="inlineStr"><is><t>South</t></is></c><c r="B3" t="inlineStr"><is><t>Accessories</t></is></c><c r="C3"><v>4</v></c><c r="D3"><v>50</v></c><c r="E3"><f>Orders[[#This Row],[Units]]*Orders[[#This Row],[Unit Price]]</f><v>200</v></c></row>',
    ).replace('</worksheet>', '<tableParts count="1"><tablePart xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></tableParts></worksheet>'),
    'xl/worksheets/_rels/sheet2.xml.rels': '<?xml version="1.0" encoding="UTF-8"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>' +
      '</Relationships>',
    'xl/tables/table1.xml': '<?xml version="1.0" encoding="UTF-8"?>' +
      '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Orders" displayName="Orders" ref="A1:E3" headerRowCount="1" totalsRowShown="0">' +
      '<autoFilter ref="A1:E3"/>' +
      '<tableColumns count="5"><tableColumn id="1" name="Region"/><tableColumn id="2" name="Category"/><tableColumn id="3" name="Units"/><tableColumn id="4" name="Unit Price"/><tableColumn id="5" name="Revenue"/></tableColumns>' +
      '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>' +
      '</table>',
    'xl/worksheets/sheet3.xml': worksheetXml(
      '<row r="1"><c r="A1" t="inlineStr"><is><t>Pivot Report</t></is></c></row>' +
      '<row r="4"><c r="A4" t="inlineStr"><is><t>Old Static Header</t></is></c></row>',
    ),
    'xl/worksheets/_rels/sheet3.xml.rels': '<?xml version="1.0" encoding="UTF-8"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>' +
      '</Relationships>',
    'xl/pivotTables/pivotTable1.xml': '<?xml version="1.0" encoding="UTF-8"?>' +
      '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="1" outline="1" outlineData="1">' +
      '<location ref="A4:D8" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>' +
      '<pivotFields count="5"><pivotField axis="axisCol" showAll="0"/><pivotField axis="axisRow" showAll="0"/><pivotField showAll="0"/><pivotField showAll="0"/><pivotField dataField="1" showAll="0"/></pivotFields>' +
      '<rowFields count="1"><field x="1"/></rowFields>' +
      '<colFields count="1"><field x="0"/></colFields>' +
      '<dataFields count="1"><dataField name="Sum of Revenue" fld="4" numFmtId="164"/></dataFields>' +
      '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0"/>' +
      '</pivotTableDefinition>',
    'xl/pivotTables/_rels/pivotTable1.xml.rels': '<?xml version="1.0" encoding="UTF-8"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/>' +
      '</Relationships>',
    'xl/pivotCache/pivotCacheDefinition1.xml': '<?xml version="1.0" encoding="UTF-8"?>' +
      '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="2">' +
      '<cacheSource type="worksheet"><worksheetSource name="Orders"/></cacheSource>' +
      '<cacheFields count="5">' +
      '<cacheField name="Region" numFmtId="0"><sharedItems count="2"><s v="North"/><s v="South"/></sharedItems></cacheField>' +
      '<cacheField name="Category" numFmtId="0"><sharedItems count="2"><s v="Accessories"/><s v="Furniture"/></sharedItems></cacheField>' +
      '<cacheField name="Units" numFmtId="0"><sharedItems/></cacheField>' +
      '<cacheField name="Unit Price" numFmtId="0"><sharedItems/></cacheField>' +
      '<cacheField name="Revenue" numFmtId="0"><sharedItems/></cacheField>' +
      '</cacheFields></pivotCacheDefinition>',
  });

  const { document, report } = await importXlsx(pkg);
  expect(report.has('tables')).toBe(false);
  expect(report.has('pivots')).toBe(false);

  // Orders table was imported with design and no external-calculation block on structured references.
  const ordersTable = document.names.find(n => n.name === 'Orders');
  expect(ordersTable?.isTable).toBe(true);
  expect(ordersTable?.area).toEqual({ sheetId: document.sheets[1]!.id, row: 1, column: 1, rows: 3, columns: 5 });
  expect(cell(document.sheets[0]!, 9, 2)?.requiresExternalCalculation).toBeUndefined();
  expect(cell(document.sheets[1]!, 2, 5)?.requiresExternalCalculation).toBeUndefined();

  // Native workbook recalculation computes structured refs & row 13 total without #NAME?.
  const recalc = runNativeGeneration(document, 'workbook');
  const dashId = document.sheets[0]!.id;
  expect(recalc.updates.find(u => u.target.sheetId === dashId && u.target.row === 9 && u.target.column === 2)?.cache)
    .toEqual({ kind: 'number', value: 100 });
  expect(recalc.updates.find(u => u.target.sheetId === dashId && u.target.row === 10 && u.target.column === 2)?.cache)
    .toEqual({ kind: 'number', value: 200 });
  expect(recalc.updates.find(u => u.target.sheetId === dashId && u.target.row === 13 && u.target.column === 2)?.cache)
    .toEqual({ kind: 'number', value: 300 });

  // Live PivotTable was materialized on Pivot sheet while preserving title cell A1.
  expect(document.pivots).toHaveLength(1);
  const pivot = document.pivots![0]!;
  expect(pivot.name).toBe('PivotTable1');
  expect(pivot.source).toEqual({ kind: 'table', tableId: ordersTable!.id });
  expect(pivot.reportLayout).toBe('compact');
  expect(pivot.values[0]?.numberFormat).toBe('$#,##0');
  const pivotSheet = document.sheets[2]!;
  expect(cell(pivotSheet, 1, 1)?.content).toEqual({ kind: 'text', value: 'Pivot Report' });
  expect(pivotSheet.cells.some(c => c.spillKind === 'pivot' && c.spillOwner === pivot.id)).toBe(true);

  const saved = await saveNativePackage(document);
  expect(await loadNativePackage(saved)).toEqual(document);
});

it('treats empty <v></v> or <v/> elements on formula and numeric cells as missing cache or blank while still rejecting non-numeric text', async () => {
  const pkg = await simplePackage(worksheetXml(
    '<row r="1"><c r="A1"><v>10</v></c><c r="A2"><v></v></c></row>' +
    '<row r="5"><c r="E5"><f>SUM(A1:A1)</f><v></v></c><c r="E6"><f>SUM(A1:A1)</f><v/></c></row>',
  ));
  const { document, report } = await importXlsx(pkg);
  const sheet = document.sheets[0]!;
  expect(cell(sheet, 1, 1)?.content).toEqual({ kind: 'number', value: 10 });
  expect(cell(sheet, 1, 2)).toBeUndefined();
  expect(cell(sheet, 5, 5)?.content).toEqual({ kind: 'formula', source: '=SUM(A1:A1)' });
  expect(cell(sheet, 5, 5)?.cache).toBeUndefined();
  expect(cell(sheet, 6, 5)?.content).toEqual({ kind: 'formula', source: '=SUM(A1:A1)' });
  expect(cell(sheet, 6, 5)?.cache).toBeUndefined();
  expect(report.has('missing-cache')).toBe(true);

  const badPkg = await simplePackage(worksheetXml('<row r="1"><c r="A1"><v>not-a-number</v></c></row>'));
  await expect(importXlsx(badPkg)).rejects.toThrow(/invalid numeric value\/cache/i);
});
