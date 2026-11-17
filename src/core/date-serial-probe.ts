/** Preserve the fictitious leap day as spreadsheet data before using a JS Date. */
export function excelDateParts(serial: number, system: 1900 | 1904 = 1900): [number, number, number] {
  const day = Math.floor(serial);
  if (system === 1900 && day === 60) return [1900, 2, 29];
  const epoch = system === 1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31);
  const date = new Date(epoch + (day - (system === 1900 && day > 60 ? 1 : 0)) * 86_400_000);
  if (!Number.isFinite(date.getTime())) throw new Error('Invalid Excel date serial');
  return [date.getUTCFullYear(), date.getUTCMonth() + 1, date.getUTCDate()];
}

export function importExcelSerial(serial: number, system: 1900 | 1904, kind: 'date' | 'time' | 'number'): number {
  return system === 1904 && kind === 'date' ? serial + 1462 : serial;
}
