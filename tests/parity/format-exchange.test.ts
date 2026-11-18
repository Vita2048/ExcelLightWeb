import { DOMParser, onErrorStopParsing } from '@xmldom/xmldom';
import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { importExcelSerial } from '../../src/core/date-serial-probe';
import { formatNumber } from '../../src/core/formats/cell-format';
import { temporalParts } from '../../src/core/formats/format-code';
import { readProbeParts } from '../../src/persistence/bounded-zip';

const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

async function readSerials(fixture: string) {
  const parts = await readProbeParts(new Blob([await readFile(new URL(fixture, import.meta.url))]));
  const parse = (path: string) => new DOMParser({ onError: onErrorStopParsing })
    .parseFromString(new TextDecoder().decode(parts.get(path)!), 'application/xml');
  const workbook = parse('xl/workbook.xml');
  const date1904 = workbook.getElementsByTagNameNS(ns, 'workbookPr')[0]?.getAttribute('date1904');
  const epoch = date1904 === '1' || date1904 === 'true' ? 1904 : 1900;
  const styles = parse('xl/styles.xml');
  const codes = new Map(Array.from(styles.getElementsByTagNameNS(ns, 'numFmt')).map(tag =>
    [tag.getAttribute('numFmtId'), tag.getAttribute('formatCode')]));
  const xfs = styles.getElementsByTagNameNS(ns, 'cellXfs')[0]!.getElementsByTagNameNS(ns, 'xf');
  const cells = parse('xl/worksheets/sheet1.xml').getElementsByTagNameNS(ns, 'c');
  return { epoch, cells: Array.from(cells).flatMap(cell => {
    const raw = cell.getElementsByTagNameNS(ns, 'v')[0]?.textContent;
    const code = codes.get(xfs[Number(cell.getAttribute('s'))]?.getAttribute('numFmtId') ?? '');
    if (raw == null || !code || ![null, 'n'].includes(cell.getAttribute('t'))) return [];
    const { date, time } = temporalParts(code);
    const serial = importExcelSerial(Number(raw), epoch, date ? 'date' : time ? 'time' : 'number');
    return [{ address: cell.getAttribute('r'), raw: Number(raw), code, serial, display: formatNumber(serial, code).text }];
  }) };
}

it('displays serial 60 from the existing desktop 1900 XLSX', async () => {
  const workbook = await readSerials('../fixtures/interchange/DesktopStyles.xlsx');
  expect(workbook.epoch).toBe(1900);
  expect(workbook.cells).toContainEqual({ address: 'A1', raw: 60, code: 'yyyy-mm-dd', serial: 60, display: '1900-02-29' });
});

it('uses the 1904 offset for date and date-time codes while keeping clocks, elapsed time and numbers', async () => {
  const workbook = await readSerials('../fixtures/formats/Desktop1904.xlsx');
  expect(workbook.epoch).toBe(1904);
  expect(workbook.cells).toEqual([
    { address: 'A1', raw: 1.5, code: '[$-405]dd.mm.yyyy', serial: 1463.5, display: '02.01.1904' },
    { address: 'A2', raw: 1.5, code: 'h:mm:ss', serial: 1.5, display: '12:00:00' },
    { address: 'A3', raw: 1.5, code: '[h]:mm:ss', serial: 1.5, display: '36:00:00' },
    { address: 'A4', raw: 1.5, code: '0.00', serial: 1.5, display: '1.50' },
    { address: 'A5', raw: 1.5, code: 'yyyy-mm-dd hh:mm', serial: 1463.5, display: '1904-01-02 12:00' },
  ]);
});
