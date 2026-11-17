import { readProbeParts, writeProbeParts } from './bounded-zip';

const guidPattern = '[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}';
const guid = new RegExp(`^${guidPattern}$`, 'iu');
const sheetPath = new RegExp(`^sheets/${guidPattern}\\.json$`, 'iu');
const utf8 = new TextDecoder('utf-8', { fatal: true });

function record(value: unknown): Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) throw new Error('Expected native JSON object');
  return value as Record<string, unknown>;
}

function id(value: unknown): string {
  if (typeof value !== 'string' || !guid.test(value) || value === '00000000-0000-0000-0000-000000000000') {
    throw new Error('Invalid native GUID');
  }
  return value.toLowerCase();
}

function parsePart(bytes: Uint8Array | undefined): Record<string, unknown> {
  if (!bytes) throw new Error('Missing native package part');
  return record(JSON.parse(utf8.decode(bytes)) as unknown);
}

export function freezeSplit(value: unknown): [number, number] {
  if (value === undefined) return [0, 0];
  if (value === 'topRow') return [1, 0];
  if (value === 'firstColumn') return [0, 1];
  const { rows, columns } = record(value);
  if (typeof rows !== 'number' || typeof columns !== 'number' || !Number.isInteger(rows) ||
      !Number.isInteger(columns) || rows < 0 || rows >= 10_000 || columns < 0 || columns >= 100 || rows + columns === 0) {
    throw new Error('Invalid frozen pane boundaries');
  }
  return [rows, columns];
}

/** Envelope proof. Full cell, style, spill, names and migration validation belongs to S17. */
export async function openNativeProbe(blob: Blob) {
  const parts = await readProbeParts(blob, name => {
    if (name !== 'manifest.json' && name !== 'workbook.json' &&
        !(name.startsWith('sheets/') && name.endsWith('.json') && sheetPath.test(name))) {
      throw new Error(`Unexpected native package entry: ${name}`);
    }
  });
  const manifest = parsePart(parts.get('manifest.json'));
  if (manifest.formatVersion !== 1 || manifest.dateSystem !== '1900') throw new Error('Unsupported native package format');
  const workbook = parsePart(parts.get('workbook.json'));
  if (id(manifest.workbookId) !== id(workbook.workbookId)) throw new Error('Workbook GUID does not match manifest');
  const order = workbook.sheets;
  if (!Array.isArray(order) || order.length < 1 || order.length > 10 || order.length !== parts.size - 2) {
    throw new Error('Sheet files do not match workbook list');
  }
  const paths = new Map([...parts.keys()].map(path => [path.toLowerCase(), path]));
  const seen = new Set<string>();
  const sheets = order.map(item => {
    const link = record(item);
    const guid = id(link.id);
    if (seen.has(guid)) throw new Error('Duplicate sheet GUID');
    seen.add(guid);
    const path = paths.get(`sheets/${guid}.json`);
    const sheet = parsePart(path === undefined ? undefined : parts.get(path));
    if (id(sheet.sheetId) !== guid || sheet.name !== link.name || typeof sheet.name !== 'string' || !Array.isArray(sheet.cells)) {
      throw new Error('Sheet contents do not match workbook list');
    }
    return { name: sheet.name, id: guid, freeze: freezeSplit(sheet.freeze), document: sheet };
  });
  return { manifest, workbook, sheets, save: () => writeProbeParts(parts) };
}
