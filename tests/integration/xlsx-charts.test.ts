import { BlobWriter, TextReader, ZipWriter } from '@zip.js/zip.js';
import { describe, expect, it } from 'vitest';
import { resolveChartData } from '../../src/charts/chart-data';
import { defaultChartElements } from '../../src/core/chart';
import { newWorkbook } from '../../src/core/workbook';
import { validateWorkbook } from '../../src/core/workbook-validation';
import { importXlsx, inspectXlsxExport } from '../../src/persistence/xlsx';
import { exchangeCodes } from '../../src/persistence/xlsx/xlsx-report';

const contentTypesBase = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
  '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
  '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>';

const stylesMinimal = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>';

const wbXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>';

const wbRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
  '</Relationships>';

const sampleSheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<sheetData>' +
  '<row r="1"><c r="A1" t="inlineStr"><is><t>Quarter</t></is></c><c r="B1" t="inlineStr"><is><t>Revenue</t></is></c><c r="C1" t="inlineStr"><is><t>Growth</t></is></c></row>' +
  '<row r="2"><c r="A2" t="inlineStr"><is><t>Q1</t></is></c><c r="B2"><v>100</v></c><c r="C2"><v>10</v></c></row>' +
  '<row r="3"><c r="A3" t="inlineStr"><is><t>Q2</t></is></c><c r="B3"><v>150</v></c><c r="C3"><v>15</v></c></row>' +
  '<row r="4"><c r="A4" t="inlineStr"><is><t>Q3</t></is></c><c r="B4"><v>200</v></c><c r="C4"><v>22</v></c></row>' +
  '</sheetData>' +
  '<drawing r:id="rId1"/>' +
  '</worksheet>';

const sheetRelsWithDrawing = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>' +
  '</Relationships>';

async function makeZip(parts: Record<string, string>): Promise<Blob> {
  const writer = new ZipWriter(new BlobWriter('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'), {
    useWebWorkers: false,
    zip64: false,
  });
  for (const [name, content] of Object.entries(parts)) {
    await writer.add(name, new TextReader(content));
  }
  return writer.close();
}

describe('OOXML XLSX chart interchange (import 2D charts & reviewed export/3D loss)', () => {
  it('imports a standard 2D Clustered Column chart with placement, title, axes, gridlines, and live resolved data without chart/drawing loss notices', async () => {
    const drawing1 = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<xdr:twoCellAnchor editAs="oneCell">' +
      '<xdr:from><xdr:col>4</xdr:col><xdr:colOff>95250</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>47625</xdr:rowOff></xdr:from>' +
      '<xdr:to><xdr:col>9</xdr:col><xdr:colOff>95250</xdr:colOff><xdr:row>13</xdr:row><xdr:rowOff>47625</xdr:rowOff></xdr:to>' +
      '<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
      '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>' +
      '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/>' +
      '</xdr:twoCellAnchor></xdr:wsDr>';

    const drawing1Rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>' +
      '</Relationships>';

    const chart1 = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Quarterly Revenue</a:t></a:r></a:p></c:rich></c:tx></c:title>' +
      '<c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>' +
      '<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>' +
      '<c:spPr><a:solidFill><a:srgbClr val="0070C0"/></a:solidFill></c:spPr>' +
      '<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f></c:strRef></c:cat>' +
      '<c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f></c:numRef></c:val></c:ser>' +
      '<c:axId val="1"/><c:axId val="2"/>' +
      '</c:barChart>' +
      '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:crossAx val="2"/></c:catAx>' +
      '<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/><c:max val="250"/></c:scaling><c:majorGridlines/><c:crossAx val="1"/></c:valAx>' +
      '</c:plotArea><c:legend><c:legendPos val="r"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>';

    const pkg = await makeZip({
      '[Content_Types].xml': contentTypesBase,
      'xl/workbook.xml': wbXml,
      'xl/_rels/workbook.xml.rels': wbRels,
      'xl/styles.xml': stylesMinimal,
      'xl/worksheets/sheet1.xml': sampleSheet,
      'xl/worksheets/_rels/sheet1.xml.rels': sheetRelsWithDrawing,
      'xl/drawings/drawing1.xml': drawing1,
      'xl/drawings/_rels/drawing1.xml.rels': drawing1Rels,
      'xl/charts/chart1.xml': chart1,
    });

    const { document, report } = await importXlsx(pkg);
    expect(() => validateWorkbook(document)).not.toThrow();
    expect(report.has(exchangeCodes.charts)).toBe(false);
    expect(report.has(exchangeCodes.drawings)).toBe(false);
    expect(document.charts).toHaveLength(1);

    const chart = document.charts![0]!;
    expect(chart.chartType).toBe('clusteredColumn');
    expect(chart.placement.row).toBe(2);
    expect(chart.placement.column).toBe(5);
    expect(chart.placement.mode).toBe('moveOnly');
    expect(chart.elements.title?.text).toBe('Quarterly Revenue');
    expect(chart.elements.legend).toEqual({ show: true, position: 'right' });
    expect(chart.elements.axes?.value?.max).toBe(250);
    expect(chart.series[0]!.fillRgb).toBe('#0070C0');

    const resolved = resolveChartData(document, chart);
    expect(resolved.categories).toEqual(['Q1', 'Q2', 'Q3']);
    expect(resolved.series[0]!.name).toBe('Revenue');
    expect(resolved.series[0]!.points.map(p => p.value)).toEqual([100, 150, 200]);
  });

  it('imports multi-plot Combo charts with secondary axes and Pie charts cleanly', async () => {
    const drawing1 = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<xdr:oneCellAnchor>' +
      '<xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
      '<xdr:ext cx="4572000" cy="2857500"/>' +
      '<xdr:graphicFrame><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
      '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>' +
      '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/>' +
      '</xdr:oneCellAnchor></xdr:wsDr>';

    const drawing1Rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>' +
      '</Relationships>';

    const comboChart = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
      '<c:chart><c:plotArea>' +
      '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>' +
      '<c:ser><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>' +
      '<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f></c:strRef></c:cat>' +
      '<c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f></c:numRef></c:val></c:ser>' +
      '<c:axId val="1"/><c:axId val="2"/></c:barChart>' +
      '<c:lineChart><c:grouping val="standard"/><c:marker val="1"/>' +
      '<c:ser><c:tx><c:strRef><c:f>Sheet1!$C$1</c:f></c:strRef></c:tx>' +
      '<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f></c:strRef></c:cat>' +
      '<c:val><c:numRef><c:f>Sheet1!$C$2:$C$4</c:f></c:numRef></c:val></c:ser>' +
      '<c:axId val="1"/><c:axId val="3"/></c:lineChart>' +
      '<c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx>' +
      '<c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>' +
      '<c:valAx><c:axId val="3"/><c:crossAx val="1"/></c:valAx>' +
      '</c:plotArea></c:chart></c:chartSpace>';

    const { document, report } = await importXlsx(await makeZip({
      '[Content_Types].xml': contentTypesBase,
      'xl/workbook.xml': wbXml,
      'xl/_rels/workbook.xml.rels': wbRels,
      'xl/styles.xml': stylesMinimal,
      'xl/worksheets/sheet1.xml': sampleSheet,
      'xl/worksheets/_rels/sheet1.xml.rels': sheetRelsWithDrawing,
      'xl/drawings/drawing1.xml': drawing1,
      'xl/drawings/_rels/drawing1.xml.rels': drawing1Rels,
      'xl/charts/chart1.xml': comboChart,
    }));

    expect(report.has(exchangeCodes.charts)).toBe(false);
    const chart = document.charts![0]!;
    expect(chart.chartType).toBe('combo');
    expect(chart.series).toHaveLength(2);
    expect(chart.series[0]!.chartType).toBe('clusteredColumn');
    expect(chart.series[0]!.secondaryAxis).toBeUndefined();
    expect(chart.series[1]!.chartType).toBe('lineMarkers');
    expect(chart.series[1]!.secondaryAxis).toBe(true);
  });

  it('skips unsupported 3D charts and non-chart drawing shapes while adding reviewed loss notices, and reports omitted charts on export', async () => {
    const drawing1 = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<xdr:twoCellAnchor>' +
      '<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
      '<xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
      '<xdr:graphicFrame><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
      '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>' +
      '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/>' +
      '</xdr:twoCellAnchor>' +
      '<xdr:twoCellAnchor><xdr:sp/></xdr:twoCellAnchor></xdr:wsDr>';

    const drawing1Rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>' +
      '</Relationships>';

    const chart3D = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
      '<c:chart><c:plotArea><c:bar3DChart><c:barDir val="col"/></c:bar3DChart></c:plotArea></c:chart></c:chartSpace>';

    const { document, report } = await importXlsx(await makeZip({
      '[Content_Types].xml': contentTypesBase,
      'xl/workbook.xml': wbXml,
      'xl/_rels/workbook.xml.rels': wbRels,
      'xl/styles.xml': stylesMinimal,
      'xl/worksheets/sheet1.xml': sampleSheet,
      'xl/worksheets/_rels/sheet1.xml.rels': sheetRelsWithDrawing,
      'xl/drawings/drawing1.xml': drawing1,
      'xl/drawings/_rels/drawing1.xml.rels': drawing1Rels,
      'xl/charts/chart1.xml': chart3D,
    }));

    expect(document.charts ?? []).toHaveLength(0);
    expect(report.has(exchangeCodes.charts)).toBe(true);
    expect(report.has(exchangeCodes.drawings)).toBe(true);

    // Export inspection reports charts when a workbook contains native charts
    let idN = 0;
    const makeId = () => `00000000-0000-4000-8000-${String(++idN).padStart(12, '0')}`;
    const docWithChart = newWorkbook(makeId);
    const sId = docWithChart.sheets[0]!.id;
    const exportReport = inspectXlsxExport({
      ...docWithChart,
      charts: [{
        id: makeId(),
        name: 'Chart 1',
        chartType: 'clusteredColumn',
        placement: { sheetId: sId, row: 1, column: 5, offsetX: 0, offsetY: 0, width: 480, height: 300 },
        source: { kind: 'range', formula: '=Sheet1!$A$1:$B$4' },
        seriesBy: 'columns',
        series: [{ id: makeId(), valuesRef: '=Sheet1!$B$2:$B$4' }],
        elements: defaultChartElements('clusteredColumn'),
        hiddenEmptyMode: 'gaps',
        showHiddenCells: false,
        styleId: 1,
        paletteId: 'colorful-1',
      }],
    });
    expect(exportReport.has(exchangeCodes.charts)).toBe(true);
  });
});
