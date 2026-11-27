import fs from 'node:fs/promises';
import path from 'node:path';
import { BlobReader, TextWriter, ZipReader } from '@zip.js/zip.js';
import { DOMParser } from '@xmldom/xmldom';

function parseArgs(argv) {
  const args = new Map();
  for (let i = 0; i < argv.length; i += 2) {
    const key = argv[i];
    const value = argv[i + 1];
    if (!key?.startsWith('--') || value === undefined) {
      throw new Error('Usage: node scripts/extract-xlsx-oracle.mjs --xlsx <file.xlsx> --spec <spec.json> --out <out.json>');
    }
    args.set(key.slice(2), value);
  }
  const xlsx = args.get('xlsx');
  const spec = args.get('spec');
  const out = args.get('out');
  if (!xlsx || !spec || !out) {
    throw new Error('Required arguments: --xlsx, --spec, --out');
  }
  return { xlsx, spec, out };
}

function columnLettersToIndex(letters) {
  let index = 0;
  for (const ch of letters.toUpperCase()) index = index * 26 + (ch.charCodeAt(0) - 64);
  return index;
}

function parseRef(ref) {
  const match = /^([A-Z]+)(\d+)$/u.exec(ref);
  if (!match) return undefined;
  return { column: columnLettersToIndex(match[1]), row: Number(match[2]) };
}

async function readZipTexts(buffer) {
  const reader = new ZipReader(new BlobReader(new Blob([buffer])));
  try {
    const entries = await reader.getEntries();
    const texts = new Map();
    for (const entry of entries) {
      if (entry.directory) continue;
      if (entry.filename === 'xl/sharedStrings.xml' || entry.filename === 'xl/worksheets/sheet1.xml') {
        texts.set(entry.filename, await entry.getData(new TextWriter()));
      }
    }
    return texts;
  } finally {
    await reader.close();
  }
}

function parseSharedStrings(xml) {
  if (!xml) return [];
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const items = [];
  const sis = doc.getElementsByTagName('si');
  for (let i = 0; i < sis.length; i++) {
    const ts = sis.item(i).getElementsByTagName('t');
    let joined = '';
    for (let j = 0; j < ts.length; j++) joined += ts.item(j).textContent ?? '';
    items.push(joined);
  }
  return items;
}

function extractSheet1Cells(xml, sharedStrings) {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const cells = new Map();
  const nodes = doc.getElementsByTagName('c');
  for (let i = 0; i < nodes.length; i++) {
    const cell = nodes.item(i);
    const ref = cell.getAttribute('r');
    const pos = ref ? parseRef(ref) : undefined;
    if (!pos) continue;
    const type = cell.getAttribute('t') ?? 'n';
    const vNode = cell.getElementsByTagName('v').item(0);
    const raw = vNode?.textContent ?? '';
    let parsed;
    if (type === 'e') parsed = { kind: 'error', value: raw.trim().toUpperCase() };
    else if (type === 'b') parsed = { kind: 'boolean', value: raw === '1' || raw.toLowerCase() === 'true' ? 'true' : 'false' };
    else if (type === 's') parsed = { kind: 'text', value: sharedStrings[Number(raw)] ?? '' };
    else if (type === 'str') parsed = { kind: 'text', value: raw };
    else if (raw === '') parsed = { kind: 'blank', value: '' };
    else parsed = { kind: 'number', value: raw };
    cells.set(`${pos.row}:${pos.column}`, parsed);
  }
  return cells;
}

async function main() {
  const { xlsx, spec: specPath, out: outPath } = parseArgs(process.argv.slice(2));
  const [xlsxBuffer, specRaw] = await Promise.all([
    fs.readFile(xlsx),
    fs.readFile(specPath, 'utf8'),
  ]);
  const spec = JSON.parse(specRaw);
  const texts = await readZipTexts(xlsxBuffer);
  const sheetXml = texts.get('xl/worksheets/sheet1.xml');
  if (!sheetXml) throw new Error('Missing xl/worksheets/sheet1.xml in XLSX package.');
  const sharedStrings = parseSharedStrings(texts.get('xl/sharedStrings.xml'));
  const cells = extractSheet1Cells(sheetXml, sharedStrings);

  // Convention: each expression in spec.expressions sits at column T (20), row (index + 1).
  const cases = spec.expressions.map((entry, index) => {
    const expression = typeof entry === 'string' ? entry : entry.expression;
    const cached = cells.get(`${index + 1}:20`) ?? { kind: 'blank', value: '' };
    return { expression, kind: cached.kind, value: cached.value };
  });

  const payload = {
    id: spec.id,
    sourceSpec: path.relative(process.cwd(), path.resolve(specPath)).replaceAll('\\', '/'),
    excelVersion: 'xlsx-cached',
    excelBuild: 'manual-save',
    capturedAt: new Date().toISOString(),
    cells: spec.cells ?? [],
    other: spec.other ?? [],
    cases,
  };
  await fs.mkdir(path.dirname(path.resolve(outPath)), { recursive: true });
  await fs.writeFile(outPath, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
  console.log(`Extracted ${cases.length} cached XLSX oracle cases -> ${outPath}`);
}

main().catch(err => {
  console.error(err);
  process.exitCode = 1;
});
