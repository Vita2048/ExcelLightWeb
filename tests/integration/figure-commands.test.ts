import { deflateSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { putFigurePlacement, removeFigurePlacement } from '../../src/application/figure-commands';
import { WorkbookSession } from '../../src/application/workbook-session';
import { emptySheet, findCell, type PythonFigure, type WorkbookDocument } from '../../src/core/workbook';
import { loadNativePackage, saveNativePackage } from '../../src/persistence/native';
import { pythonBlockFromOutput } from '../../src/python/python-results';

function crc32(bytes: Uint8Array) {
  let value = 0xffffffff;
  for (const byte of bytes) {
    value ^= byte;
    for (let bit = 0; bit < 8; bit++) value = (value & 1) ? (value >>> 1) ^ 0xedb88320 : value >>> 1;
  }
  return (value ^ 0xffffffff) >>> 0;
}

function uint32(value: number) {
  return Uint8Array.of(value >>> 24, value >>> 16, value >>> 8, value);
}

function pngChunk(name: string, payload: Uint8Array) {
  const type = new TextEncoder().encode(name);
  const body = new Uint8Array(type.length + payload.length);
  body.set(type); body.set(payload, type.length);
  const result = new Uint8Array(12 + payload.length);
  result.set(uint32(payload.length)); result.set(body, 4); result.set(uint32(crc32(body)), 8 + payload.length);
  return result;
}

function onePixelPng(): string {
  const signature = Uint8Array.of(137, 80, 78, 71, 13, 10, 26, 10);
  const ihdr = Uint8Array.of(0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0);
  const pixel = new Uint8Array(deflateSync(Uint8Array.of(0, 24, 96, 160, 255)));
  const chunks = [pngChunk('IHDR', ihdr), pngChunk('IDAT', pixel), pngChunk('IEND', new Uint8Array())];
  const result = new Uint8Array(signature.length + chunks.reduce((sum, chunk) => sum + chunk.length, 0));
  let offset = 0;
  for (const part of [signature, ...chunks]) { result.set(part, offset); offset += part.length; }
  return Buffer.from(result).toString('base64');
}

function fixture() {
  let next = 5;
  const makeId = () => `00000000-0000-4000-8000-${String(++next).padStart(12, '0')}`;
  const [bookId, sourceId, chartsId, cellId, blockId] = [1, 2, 3, 4, 5].map(index =>
    `00000000-0000-4000-8000-${String(index).padStart(12, '0')}`);
  const figure: PythonFigure = { pngBase64: onePixelPng(), width: 1, height: 1 };
  const source = emptySheet(sourceId!, 'Source');
  const charts = emptySheet(chartsId!, 'Charts');
  const document: WorkbookDocument = { id: bookId!, formatVersion: 1, dateSystem: '1900', names: [], sheets: [
    { ...source, cells: [{ id: cellId!, row: 1, column: 4, content: { kind: 'python', blockId: blockId!, source: 'plot()',
      block: { id: blockId!, source: 'plot()', bindings: [], lastValue: null, hasResult: true,
        inputFingerprint: null, runtimeFingerprint: null, status: 'succeeded', error: null,
        lastTable: null, spillActive: false, lastFigure: figure, placements: [] } } }] },
    charts,
  ] };
  const session = new WorkbookSession(document, { makeId });
  const placement = (id: string, width: number, height: number) => ({ id, sheetId: chartsId!, row: 2, column: 2,
    rows: 9, columns: 5, offsetX: 8, offsetY: 5, width, height,
    renderedFigure: figure, renderedSourceId: 'A'.repeat(64) });
  return { session, figure, placement, blockId: blockId!, chartsId: chartsId! };
}

describe('S23 Figure placement session commands', () => {
  it('keeps the live worker cache key out of workbook Figure DTOs', () => {
    const { session } = fixture();
    const cell = findCell(session.active.sheets[0]!, { row: 1, column: 4 })!;
    if (cell.content.kind !== 'python' || !cell.content.block) throw new Error('Python Figure source missing.');
    const output = pythonBlockFromOutput(cell.content.block, { status: 'succeeded', result: { kind: 'figure', figure: {
      pngBase64: cell.content.block.lastFigure!.pngBase64, width: 1, height: 1, figureId: 'B'.repeat(64),
    } }, stdout: '', stderr: '', outputTruncated: false }, 'input', 'runtime');
    expect(output.lastFigure).toEqual({ pngBase64: cell.content.block.lastFigure!.pngBase64, width: 1, height: 1 });
  });

  it('keeps multiple independent sizes, supports geometry undo/redo, and removes one placement', () => {
    const { session, placement, blockId } = fixture();
    const first = placement('00000000-0000-4000-8000-000000000006', 480, 240);
    const second = placement('00000000-0000-4000-8000-000000000007', 300, 180);
    let document = putFigurePlacement(session.active, blockId, first);
    document = putFigurePlacement(document, blockId, second);
    session.execute({ kind: 'structure', document, label: 'Place two Figures' });
    let block = findCell(session.active.sheets[0]!, { row: 1, column: 4 })!.content;
    expect(block.kind === 'python' && block.block?.placements).toMatchObject([
      { id: first.id, width: 480, height: 240, renderedSourceId: 'A'.repeat(64) },
      { id: second.id, width: 300, height: 180, renderedSourceId: 'A'.repeat(64) },
    ]);

    const moved = putFigurePlacement(session.active, blockId, { ...first, row: 8, column: 7, offsetX: 22, width: 540 });
    session.execute({ kind: 'structure', document: moved, label: 'Move or resize Figure' });
    expect(session.undo()).toBe(true);
    block = findCell(session.active.sheets[0]!, { row: 1, column: 4 })!.content;
    expect(block.kind === 'python' && block.block?.placements?.[0]).toMatchObject({ row: 2, column: 2, width: 480 });
    expect(session.redo()).toBe(true);
    block = findCell(session.active.sheets[0]!, { row: 1, column: 4 })!.content;
    expect(block.kind === 'python' && block.block?.placements?.[0]).toMatchObject({ row: 8, column: 7, width: 540 });

    const deleted = removeFigurePlacement(session.active, first.id);
    session.execute({ kind: 'structure', document: deleted, label: 'Delete Figure placement' });
    block = findCell(session.active.sheets[0]!, { row: 1, column: 4 })!.content;
    expect(block.kind === 'python' && block.block?.placements?.map(item => item.id)).toEqual([second.id]);
  });

  it('rejects a late placement proposal after the workbook basis changed', () => {
    const { session, placement, blockId, chartsId } = fixture();
    const basis = session.capture();
    const proposal = putFigurePlacement(basis.document, blockId, placement('00000000-0000-4000-8000-000000000006', 480, 240));
    session.execute({ kind: 'set-cells', edits: [{ target: { sheetId: chartsId, row: 1, column: 1 },
      content: { kind: 'text', value: 'changed while redrawing' } }] });
    expect(session.tryExecute(basis, { kind: 'structure', label: 'Late Figure placement', document: proposal })).toBe(false);
    const block = findCell(session.active.sheets[0]!, { row: 1, column: 4 })!.content;
    expect(block.kind === 'python' && block.block?.placements).toEqual([]);
  });

  it('saves/reopens placement geometry and PNG while leaving live worker identity session-only', async () => {
    const { session, figure, placement, blockId, chartsId } = fixture();
    const document = putFigurePlacement(session.active, blockId, placement('00000000-0000-4000-8000-000000000006', 480, 240));
    session.execute({ kind: 'structure', document, label: 'Place Figure' });
    const reopened = await loadNativePackage(await saveNativePackage(session.active));
    const cell = findCell(reopened.sheets[0]!, { row: 1, column: 4 })!;
    if (cell.content.kind !== 'python' || !cell.content.block) throw new Error('Python Figure source missing after reopen.');
    expect(cell.content.block.lastFigure).toEqual({ pngBase64: figure.pngBase64, width: 1, height: 1 });
    expect(cell.content.block.placements).toMatchObject([{ id: '00000000-0000-4000-8000-000000000006', sheetId: chartsId, row: 2,
      column: 2, rows: 9, columns: 5, offsetX: 8, offsetY: 5, width: 480, height: 240,
      renderedFigure: { pngBase64: figure.pngBase64, width: 1, height: 1 }, renderedSourceId: 'A'.repeat(64) }]);
  });
});
