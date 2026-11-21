import { execFileSync } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { BlobWriter, TextReader, ZipWriter } from '@zip.js/zip.js';
import { WorkbookSession } from '../../src/application/workbook-session';
import { ownDto, type CellDocument, type CellStyle, type WorkbookDocument } from '../../src/core/workbook';
import { equalDto } from '../../src/core/workbook-delta';
import {
  defaultPackageLimits,
  loadNativePackage,
  NativePackage,
  PackageException,
  saveNativePackage,
  SaveFault,
} from '../../src/persistence/native';

const readFixture = async (name: string) => new Blob([await readFile(new URL(`../fixtures/${name}`, import.meta.url))]);
const runDesktopInterop = (args: string[], timeout = 30000) => execFileSync('dotnet', args, {
  encoding: 'utf-8',
  timeout,
  env: { ...process.env, DOTNET_CLI_TELEMETRY_OPTOUT: '1', DOTNET_NOLOGO: '1' },
});

function storedZip(entries: readonly (readonly [string, string])[]): Blob {
  const encoder = new TextEncoder();
  const locals: Uint8Array[] = [];
  const centrals: Uint8Array[] = [];
  let localOffset = 0;
  let centralSize = 0;
  const crc32 = (bytes: Uint8Array) => {
    let value = 0xffffffff;
    for (const byte of bytes) {
      value ^= byte;
      for (let bit = 0; bit < 8; bit++) value = (value & 1) !== 0 ? (value >>> 1) ^ 0xedb88320 : value >>> 1;
    }
    return (value ^ 0xffffffff) >>> 0;
  };
  const write16 = (view: DataView, offset: number, value: number) => view.setUint16(offset, value, true);
  const write32 = (view: DataView, offset: number, value: number) => view.setUint32(offset, value, true);

  for (const [filename, text] of entries) {
    const name = encoder.encode(filename);
    const data = encoder.encode(text);
    const crc = crc32(data);
    const local = new Uint8Array(30 + name.length + data.length);
    const localView = new DataView(local.buffer);
    write32(localView, 0, 0x04034b50);
    write16(localView, 4, 20);
    write32(localView, 14, crc);
    write32(localView, 18, data.length);
    write32(localView, 22, data.length);
    write16(localView, 26, name.length);
    local.set(name, 30);
    local.set(data, 30 + name.length);
    locals.push(local);

    const central = new Uint8Array(46 + name.length);
    const centralView = new DataView(central.buffer);
    write32(centralView, 0, 0x02014b50);
    write16(centralView, 4, 20);
    write16(centralView, 6, 20);
    write32(centralView, 16, crc);
    write32(centralView, 20, data.length);
    write32(centralView, 24, data.length);
    write16(centralView, 28, name.length);
    write32(centralView, 42, localOffset);
    central.set(name, 46);
    centrals.push(central);

    localOffset += local.length;
    centralSize += central.length;
  }

  const end = new Uint8Array(22);
  const endView = new DataView(end.buffer);
  write32(endView, 0, 0x06054b50);
  write16(endView, 8, entries.length);
  write16(endView, 10, entries.length);
  write32(endView, 12, centralSize);
  write32(endView, 16, localOffset);
  const parts: BlobPart[] = [];
  for (const bytes of [...locals, ...centrals, end]) {
    const copy = new Uint8Array(bytes.length);
    copy.set(bytes);
    parts.push(copy.buffer);
  }
  return new Blob(parts, { type: 'application/zip' });
}

describe('NativePackage round-trip and desktop equivalence', () => {
  it.each([
    'interchange/DesktopStyles.exlt',
    'sales/Showcase.exlt',
    'sla/Support-SLA.exlt',
  ])('round-trips all JSON parts, styles, and Python metadata from %s', async fixtureName => {
    const originalBlob = await readFixture(fixtureName);
    const doc = await NativePackage.load(originalBlob);

    expect(doc.formatVersion).toBe(1);
    expect(doc.dateSystem).toBe('1900');
    expect(doc.sheets.length).toBeGreaterThan(0);

    // Save using native writer
    const savedBlob = await NativePackage.save(doc);
    expect(savedBlob.size).toBeGreaterThan(0);

    // Re-load and verify exact equality
    const reloaded = await NativePackage.load(savedBlob);
    expect(equalDto(doc, reloaded)).toBe(true);
  });

  it('proves desktop C# NativePackage.Load accepts web-saved DesktopStyles.exlt', async () => {
    const originalBlob = await readFixture('interchange/DesktopStyles.exlt');
    const doc = await loadNativePackage(originalBlob);
    const savedBlob = await saveNativePackage(doc);

    const tempDir = await mkdtemp(join(tmpdir(), 'exlt-interchange-'));
    try {
      const bytes = new Uint8Array(await savedBlob.arrayBuffer());
      await writeFile(join(tempDir, 'DesktopStyles.exlt'), bytes);

      // Also copy DesktopStyles.xlsx for the full interchange check
      const xlsxBlob = await readFixture('interchange/DesktopStyles.xlsx');
      await writeFile(join(tempDir, 'DesktopStyles.xlsx'), new Uint8Array(await xlsxBlob.arrayBuffer()));

      const output = runDesktopInterop([
        'run',
        '--project',
        'tests/desktop-interchange/Interchange.csproj',
        '--',
        'check',
        tempDir,
      ]);

      expect(output).toContain('Desktop native equality and XLSX style reimport passed.');
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  }, 30000);

  it('compares every native field with the desktop after web round-trips of representative fixtures', async () => {
    const fixtureNames = [
      'interchange/DesktopStyles.exlt',
      'sales/Showcase.exlt',
      'sla/Support-SLA.exlt',
    ];
    const tempDir = await mkdtemp(join(tmpdir(), 'exlt-desktop-equality-'));
    try {
      const args = [
        'run',
        '--project',
        'tests/desktop-interchange/Interchange.csproj',
        '--',
        'compare',
      ];
      for (const [index, fixtureName] of fixtureNames.entries()) {
        const originalBlob = await readFixture(fixtureName);
        const original = await loadNativePackage(originalBlob);
        const rewritten = await saveNativePackage(original);
        const originalPath = join(tempDir, `original-${index}.exlt`);
        const rewrittenPath = join(tempDir, `rewritten-${index}.exlt`);
        await writeFile(originalPath, new Uint8Array(await originalBlob.arrayBuffer()));
        await writeFile(rewrittenPath, new Uint8Array(await rewritten.arrayBuffer()));
        args.push(originalPath, rewrittenPath);
      }

      const output = runDesktopInterop(args);
      expect(output).toContain('Desktop native round-trip equality passed for 3 fixtures.');
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  }, 30000);

  it('keeps formula source and its derived cache distinct in the native package', async () => {
    const original = await loadNativePackage(await readFixture('sales/Showcase.exlt'));
    const formula = original.sheets.flatMap(sheet => sheet.cells)
      .find(cell => cell.content.kind === 'formula' && cell.cache !== undefined);
    expect(formula?.content.kind).toBe('formula');
    expect(formula?.cache).toBeDefined();
    if (!formula || formula.content.kind !== 'formula') return;

    const reloaded = await loadNativePackage(await saveNativePackage(original));
    const copy = reloaded.sheets.flatMap(sheet => sheet.cells).find(cell => cell.id === formula.id);
    expect(copy?.content).toEqual(formula.content);
    expect(copy?.cache).toEqual(formula.cache);
    expect(copy?.content.kind).toBe('formula');
  });

  it('preserves empty text, caches, styles, and distinct missing cells', async () => {
    const originalBlob = await readFixture('interchange/DesktopStyles.exlt');
    const doc = await loadNativePackage(originalBlob);
    const sheet = doc.sheets[0]!;

    expect(sheet.cells.length).toBe(2);
    expect(sheet.cells[0]!.content).toEqual({ kind: 'number', value: 60 });
    expect(sheet.cells[0]!.style?.numberFormat).toBe('yyyy-mm-dd');
    expect(sheet.cells[1]!.content).toEqual({ kind: 'text', value: 'stacked' });
    expect(sheet.frozenPanes).toEqual({ rows: 1, columns: 0 });
  });

  it('migrates legacy detached table headers on package load', async () => {
    // Create an archive with an unmigrated table (headerRow = false, no headerLabels)
    const writer = new ZipWriter(new BlobWriter('application/zip'), { useWebWorkers: false });
    const wbId = '11111111-1111-1111-1111-111111111111';
    const sheetId = '22222222-2222-2222-2222-222222222222';
    const tableId = '33333333-3333-3333-3333-333333333333';

    await writer.add('manifest.json', new TextReader(JSON.stringify({
      formatVersion: 1,
      workbookId: wbId,
      dateSystem: '1900',
    })));

    await writer.add('workbook.json', new TextReader(JSON.stringify({
      workbookId: wbId,
      names: [{
        Id: tableId,
        Name: 'LegacyTable',
        Area: { SheetId: sheetId, Row: 1, Column: 1, Rows: 3, Columns: 2 },
        IsTable: true,
        Design: { Style: 'Blue', HeaderRow: false, TotalRow: false },
        HeaderLabels: [],
      }],
      sheets: [{ id: sheetId, name: 'Sheet1' }],
    })));

    await writer.add(`sheets/${sheetId}.json`, new TextReader(JSON.stringify({
      sheetId,
      name: 'Sheet1',
      cells: [
        { id: '44444444-4444-4444-4444-444444444444', row: 1, column: 1, content: { kind: 'text', value: 'ColA' } },
        { id: '55555555-5555-5555-5555-555555555555', row: 1, column: 2, content: { kind: 'text', value: 'ColB' } },
        { id: '66666666-6666-6666-6666-666666666666', row: 2, column: 1, content: { kind: 'number', value: 10 } },
        { id: '77777777-7777-7777-7777-777777777777', row: 2, column: 2, content: { kind: 'number', value: 20 } },
      ],
    })));

    const blob = await writer.close();
    const doc = await loadNativePackage(blob);

    const table = doc.names.find(n => n.name === 'LegacyTable')!;
    expect(table).toBeDefined();
    // Migrated: first row shifted down, headerLabels captured
    expect(table.area.row).toBe(2);
    expect(table.area.rows).toBe(2);
    expect(table.headerLabels).toEqual(['ColA', 'ColB']);
  });
});

describe('NativePackage hostile rejection and budget enforcement', () => {
  async function createZip(entries: Record<string, string> | readonly (readonly [string, string])[]): Promise<Blob> {
    const writer = new ZipWriter(new BlobWriter('application/zip'), { useWebWorkers: false });
    const pairs: readonly (readonly [string, string])[] = Array.isArray(entries) ? entries : Object.entries(entries);
    for (const [name, content] of pairs) {
      await writer.add(name, new TextReader(content));
    }
    return writer.close();
  }

  const validManifest = JSON.stringify({
    formatVersion: 1,
    workbookId: '11111111-1111-1111-1111-111111111111',
    dateSystem: '1900',
  });

  const validWorkbook = JSON.stringify({
    workbookId: '11111111-1111-1111-1111-111111111111',
    sheets: [{ id: '22222222-2222-2222-2222-222222222222', name: 'Sheet1' }],
  });

  const validSheet = JSON.stringify({
    sheetId: '22222222-2222-2222-2222-222222222222',
    name: 'Sheet1',
    cells: [],
  });

  it('rejects unsupported format version', async () => {
    const blob = await createZip({
      'manifest.json': JSON.stringify({ formatVersion: 2, workbookId: '11111111-1111-1111-1111-111111111111', dateSystem: '1900' }),
      'workbook.json': validWorkbook,
      'sheets/22222222-2222-2222-2222-222222222222.json': validSheet,
    });
    await expect(loadNativePackage(blob)).rejects.toThrow('Unsupported package version 2. This build reads version 1.');
  });

  it('rejects non-1900 date system', async () => {
    const blob = await createZip({
      'manifest.json': JSON.stringify({ formatVersion: 1, workbookId: '11111111-1111-1111-1111-111111111111', dateSystem: '1904' }),
      'workbook.json': validWorkbook,
      'sheets/22222222-2222-2222-2222-222222222222.json': validSheet,
    });
    await expect(loadNativePackage(blob)).rejects.toThrow("Date system '1904' is not supported. This prototype does not convert other epochs.");
  });

  it('rejects mismatched workbook GUID', async () => {
    const blob = await createZip({
      'manifest.json': validManifest,
      'workbook.json': JSON.stringify({
        workbookId: '99999999-9999-9999-9999-999999999999',
        sheets: [{ id: '22222222-2222-2222-2222-222222222222', name: 'Sheet1' }],
      }),
      'sheets/22222222-2222-2222-2222-222222222222.json': validSheet,
    });
    await expect(loadNativePackage(blob)).rejects.toThrow('Workbook id does not match the manifest.');
  });

  it('rejects path traversal in entry name', async () => {
    const blob = await createZip({
      'manifest.json': validManifest,
      'workbook.json': validWorkbook,
      '../outside.json': '{}',
    });
    await expect(loadNativePackage(blob)).rejects.toThrow('Package entry path is not allowed: ../outside.json');
  });

  it('rejects executable metadata in entry path', async () => {
    const blob = await createZip({
      'manifest.json': validManifest,
      'workbook.json': validWorkbook,
      'sheets/22222222-2222-2222-2222-222222222222.json': validSheet,
      'python/block.json': '{"source":"print(1)"}',
    });
    await expect(loadNativePackage(blob)).rejects.toThrow('Executable metadata is not accepted in package version 1.');
  });

  it('rejects unexpected package entries', async () => {
    const blob = await createZip({
      'manifest.json': validManifest,
      'workbook.json': validWorkbook,
      'sheets/22222222-2222-2222-2222-222222222222.json': validSheet,
      'extra.json': '{}',
    });
    await expect(loadNativePackage(blob)).rejects.toThrow("Unexpected package entry 'extra.json'.");
  });

  it('rejects duplicate JSON property in manifest', async () => {
    const blob = await createZip({
      'manifest.json': '{"formatVersion":1,"formatVersion":1,"workbookId":"11111111-1111-1111-1111-111111111111","dateSystem":"1900"}',
      'workbook.json': validWorkbook,
      'sheets/22222222-2222-2222-2222-222222222222.json': validSheet,
    });
    await expect(loadNativePackage(blob)).rejects.toThrow('Duplicate JSON property.');
  });

  it('rejects duplicate ZIP entry names', async () => {
    const blob = storedZip([
      ['manifest.json', validManifest],
      ['manifest.json', validManifest],
      ['workbook.json', validWorkbook],
      ['sheets/22222222-2222-2222-2222-222222222222.json', validSheet],
    ]);
    await expect(loadNativePackage(blob)).rejects.toThrow("Duplicate package entry 'manifest.json'.");
  });

  it('rejects malformed package JSON', async () => {
    const blob = await createZip({
      'manifest.json': '{',
      'workbook.json': validWorkbook,
      'sheets/22222222-2222-2222-2222-222222222222.json': validSheet,
    });
    await expect(loadNativePackage(blob)).rejects.toThrow('Malformed package JSON.');
  });

  it('rejects JSON depth exceeding 16', async () => {
    const deepObj = '{"a":{"a":{"a":{"a":{"a":{"a":{"a":{"a":{"a":{"a":{"a":{"a":{"a":{"a":{"a":{"a":{"a":1}}}}}}}}}}}}}}}}}';
    const blob = await createZip({
      'manifest.json': validManifest,
      'workbook.json': deepObj,
      'sheets/22222222-2222-2222-2222-222222222222.json': validSheet,
    });
    await expect(loadNativePackage(blob)).rejects.toThrow('Malformed package JSON.');
  });

  it('rejects oversize package', async () => {
    const blob = await createZip({
      'manifest.json': validManifest,
      'workbook.json': validWorkbook,
      'sheets/22222222-2222-2222-2222-222222222222.json': validSheet,
    });
    await expect(loadNativePackage(blob, { maxPackageBytes: 10 })).rejects.toThrow('Package exceeds the 10 byte limit.');
  });

  it('rejects oversize entry', async () => {
    const blob = await createZip({
      'manifest.json': validManifest,
      'workbook.json': validWorkbook,
      'sheets/22222222-2222-2222-2222-222222222222.json': validSheet,
    });
    await expect(loadNativePackage(blob, { maxEntryBytes: 10 })).rejects.toThrow('exceeds the entry limit.');
  });

  it('enforces expanded byte and entry-count limits while streaming', async () => {
    const blob = await createZip({
      'manifest.json': validManifest,
      'workbook.json': validWorkbook,
      'sheets/22222222-2222-2222-2222-222222222222.json': validSheet,
    });
    await expect(loadNativePackage(blob, { maxExpandedBytes: 32 })).rejects.toThrow('exceeds the size limit.');
    await expect(loadNativePackage(blob, { maxEntries: 2 })).rejects.toThrow('Package has more than 2 entries.');
  });

  it('rejects Python bindings outside the desktop A1 reference contract', async () => {
    const pythonSheet = JSON.stringify({
      sheetId: '22222222-2222-2222-2222-222222222222',
      name: 'Sheet1',
      cells: [{
        id: '33333333-3333-3333-3333-333333333333', row: 1, column: 1,
        content: { kind: 'python', block: {
          Id: '44444444-4444-4444-4444-444444444444', Source: '',
          Bindings: [{ Name: 'data', Reference: 'not a cell reference' }],
        } },
      }],
    });
    const blob = await createZip({
      'manifest.json': validManifest,
      'workbook.json': validWorkbook,
      'sheets/22222222-2222-2222-2222-222222222222.json': pythonSheet,
    });
    await expect(loadNativePackage(blob)).rejects.toThrow('Bindings need a unique name and one A1 cell/range reference.');
  });

  it('rejects Python table error caches and invalid figure payloads', async () => {
    const makePythonPackage = (block: Record<string, unknown>) => createZip({
      'manifest.json': validManifest,
      'workbook.json': validWorkbook,
      'sheets/22222222-2222-2222-2222-222222222222.json': JSON.stringify({
        sheetId: '22222222-2222-2222-2222-222222222222',
        name: 'Sheet1',
        cells: [{
          id: '33333333-3333-3333-3333-333333333333', row: 1, column: 1,
          content: { kind: 'python', block },
        }],
      }),
    });
    const tableBlob = await makePythonPackage({
      Id: '44444444-4444-4444-4444-444444444444', Source: '', Bindings: [],
      LastValue: { Kind: 'error', Text: '#VALUE!' }, HasResult: true,
      LastTable: { Rows: 1, Columns: 1, Cells: [{ Kind: 'error', Text: '#VALUE!' }], Notices: [] },
    });
    await expect(loadNativePackage(tableBlob)).rejects.toThrow('Unsupported table value.');

    const figureBlob = await makePythonPackage({
      Id: '44444444-4444-4444-4444-444444444444', Source: '', Bindings: [],
      HasResult: true, LastFigure: { PngBase64: 'not base64', Width: 1, Height: 1 },
    });
    await expect(loadNativePackage(figureBlob)).rejects.toThrow('Invalid figure PNG encoding.');
  });

  it('rejects truncated archive', async () => {
    const truncated = new Blob([new Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2, 3, 4])]);
    await expect(loadNativePackage(truncated)).rejects.toThrow('Package is not a valid .exlt archive.');
  });

  it('rejects formulas without leading =', async () => {
    const badSheet = JSON.stringify({
      sheetId: '22222222-2222-2222-2222-222222222222',
      name: 'Sheet1',
      cells: [{
        id: '33333333-3333-3333-3333-333333333333',
        row: 1,
        column: 1,
        content: { kind: 'formula', source: 'SUM(1)' },
      }],
    });
    const blob = await createZip({
      'manifest.json': validManifest,
      'workbook.json': validWorkbook,
      'sheets/22222222-2222-2222-2222-222222222222.json': badSheet,
    });
    await expect(loadNativePackage(blob)).rejects.toThrow('A formula must start with = and have valid source text.');
  });

  it('rejects non-finite number values', async () => {
    const badSheet = JSON.stringify({
      sheetId: '22222222-2222-2222-2222-222222222222',
      name: 'Sheet1',
      cells: [{
        id: '33333333-3333-3333-3333-333333333333',
        row: 1,
        column: 1,
        content: { kind: 'number', value: 1e9999 },
      }],
    });
    const blob = await createZip({
      'manifest.json': validManifest,
      'workbook.json': validWorkbook,
      'sheets/22222222-2222-2222-2222-222222222222.json': badSheet,
    });
    await expect(loadNativePackage(blob)).rejects.toThrow('value must be a finite number.');
  });

  it('rejects sheet file count mismatch', async () => {
    const blob = await createZip({
      'manifest.json': validManifest,
      'workbook.json': validWorkbook,
      // no sheet file!
    });
    await expect(loadNativePackage(blob)).rejects.toThrow('Sheet files do not match the workbook sheet list.');
  });
});

describe('WorkbookSession adoption and save safety', () => {
  it('preserves existing session document on failed package parse', async () => {
    const initialBlob = await readFixture('interchange/DesktopStyles.exlt');
    const doc = await loadNativePackage(initialBlob);

    const session = new WorkbookSession(doc, { makeId: () => crypto.randomUUID() });
    const beforeCapture = session.capture();

    // Parse fully before adoption: a failed decode cannot reach session.adopt.
    const badBlob = new Blob([new Uint8Array([1, 2, 3, 4])]);
    await expect((async () => {
      const candidate = await loadNativePackage(badBlob);
      session.adopt(candidate);
    })()).rejects.toThrow(PackageException);

    // Active document is completely untouched
    expect(session.capture()).toBe(beforeCapture);
    expect(equalDto(session.active, doc)).toBe(true);
    expect(session.isDirty).toBe(false);
  });

  it('enforces atomic save safety with fault injection', async () => {
    const doc = await loadNativePackage(await readFixture('interchange/DesktopStyles.exlt'));

    await expect(saveNativePackage(doc, { fault: SaveFault.DuringWrite })).rejects.toThrow('Injected write failure.');

    let replaced = false;
    await expect(saveNativePackage(doc, {
      beforeReplace: () => { replaced = true; },
      fault: SaveFault.BeforeReplace,
    })).rejects.toThrow('Injected replace failure.');
    expect(replaced).toBe(true);
  });
});
