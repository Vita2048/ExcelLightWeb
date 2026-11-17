import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { expect, it } from 'vitest';
import { BlobWriter, Uint8ArrayReader, ZipWriter } from '@zip.js/zip.js';
import { ExpansionBudget, probeZipBudget, readProbeParts, visitZip, writeProbeParts } from '../../src/persistence/bounded-zip';
import { freezeSplit, openNativeProbe } from '../../src/persistence/native-probe';
import { rewriteXlsxProbe, roundTripStyles } from '../../src/persistence/xlsx-style-probe';

const readFixture = async (name: string) => new Blob([await readFile(new URL(`../fixtures/${name}`, import.meta.url))]);

it.each(['sales/Showcase.exlt', 'sla/Support-SLA.exlt', 'interchange/DesktopStyles.exlt'])(
  'round-trips all JSON parts and identities from %s', async name => {
    const original = await readFixture(name);
    const workbook = await openNativeProbe(original);
    const repacked = await workbook.save();
    const reopen = await openNativeProbe(repacked);
    expect(reopen.manifest).toEqual(workbook.manifest);
    expect(reopen.sheets).toEqual(workbook.sheets);
    const oldParts = await readProbeParts(original);
    const newParts = await readProbeParts(repacked);
    expect([...newParts.keys()]).toEqual([...oldParts.keys()]);
    for (const [key, bytes] of oldParts) {
      const next = newParts.get(key);
      expect(next?.length, key).toBe(bytes.length);
      expect(next?.every((byte, i) => byte === bytes[i]), key).toBe(true);
    }
    const output = process.env.S03_INTERCHANGE_OUT;
    if (output && name.endsWith('DesktopStyles.exlt')) {
      await mkdir(output, { recursive: true });
      await writeFile(join(output, 'DesktopStyles.exlt'), new Uint8Array(await repacked.arrayBuffer()));
    }
  },
);

it('reads legacy and extended freeze fields and rejects invalid coordinates', async () => {
  const workbook = await openNativeProbe(await readFixture('interchange/DesktopStyles.exlt'));
  expect(workbook.sheets[0]).toMatchObject({ id: 'b71420c7-ad56-4000-8000-000000000001', freeze: [1, 0] });
  expect(freezeSplit('firstColumn')).toEqual([0, 1]);
  expect(freezeSplit({ rows: 2, columns: 3 })).toEqual([2, 3]);
  expect(() => freezeSplit({ rows: 10_000, columns: 1 })).toThrow('frozen pane');
  expect(() => freezeSplit('TopRow')).toThrow();
});

it('rejects future native versions and mismatched workbook GUIDs', async () => {
  const parts = await readProbeParts(await readFixture('interchange/DesktopStyles.exlt'));
  const original = new TextDecoder().decode(parts.get('manifest.json'));
  const manifest = JSON.parse(original);
  const setManifest = (value: object) => parts.set('manifest.json', new TextEncoder().encode(JSON.stringify(value)));
  setManifest({ ...manifest, formatVersion: 2 });
  await expect(openNativeProbe(await writeProbeParts(parts))).rejects.toThrow('Unsupported native package format');
  setManifest({ ...manifest, workbookId: 'b71420c7-ad56-4000-8000-000000000009' });
  await expect(openNativeProbe(await writeProbeParts(parts))).rejects.toThrow('GUID does not match');
});

it('checks streamed output against remaining per-entry and total budget', () => {
  const limits = { ...probeZipBudget, entryBytes: 8, expandedBytes: 10 };
  const budget = new ExpansionBudget(limits);
  const first = budget.begin('first', 2);
  first(5);
  expect(() => first(4)).toThrow('expanded byte limit');
  const second = budget.begin('second', 2);
  second(4);
  expect(() => second(2)).toThrow('expanded byte limit');
});

it('enforces the byte budget on real decoded data even with understated ZIP metadata', async () => {
  const writer = new ZipWriter(new BlobWriter(), { useWebWorkers: false, zip64: false });
  await writer.add('entry.json', new Uint8ArrayReader(new Uint8Array(8192).fill(32)), { dataDescriptor: false });
  const bytes = await (await writer.close()).arrayBuffer();
  const view = new DataView(bytes);
  const directory = view.getUint32(bytes.byteLength - 6, true);
  expect(view.getUint32(directory, true)).toBe(0x02014b50);
  view.setUint32(22, 1, true);
  view.setUint32(directory + 24, 1, true);
  let delivered = 0;
  let finished = false;
  const limits = { ...probeZipBudget, entryBytes: 64, expandedBytes: 64 };
  await expect(visitZip(new Blob([bytes]), limits, () => ({
    write(chunk) { delivered += chunk.length; },
    finish() { finished = true; },
  }))).rejects.toThrow(/expanded byte limit|uncompressed size/i);
  expect(delivered).toBeLessThanOrEqual(64);
  expect(finished).toBe(false);
});

it('preserves extended styles from a desktop XLSX across ZIP and OOXML serialization', async () => {
  const fixture = await readFixture('interchange/DesktopStyles.xlsx');
  const output = await rewriteXlsxProbe(fixture);
  expect(output.borders).toContainEqual({
    up: true, down: true,
    pens: {
      top: { style: 'slantDashDot', color: 'FF112233' }, bottom: { style: 'mediumDashDotDot', color: 'FF445566' },
      right: { style: 'double', color: null }, diagonal: { style: 'dotted', color: 'FF778899' },
    },
  });
  expect(output.gradients).toEqual([
    { geometry: { type: 'linear', degree: 225 }, stops: [
      { position: 0, argb: 'FFFFFFFF' }, { position: 0.5, argb: 'FF2288AA' }, { position: 1, argb: 'FFFFFFFF' },
    ] },
    { geometry: { type: 'path', offsets: [0.25, 0.25, 0.25, 0.25] }, stops: [
      { position: 0, argb: 'FFFFFFFF' }, { position: 1, argb: 'FF2288AA' },
    ] },
  ]);
  expect(output.rotations).toEqual([0, 0, -45, 'stacked', -45, 'stacked']);
  const reopened = await rewriteXlsxProbe(output.blob);
  expect(reopened).toMatchObject({ borders: output.borders, gradients: output.gradients, rotations: output.rotations });
  const before = await readProbeParts(fixture);
  const after = await readProbeParts(output.blob);
  for (const [name, data] of before) {
    if (name !== 'xl/styles.xml') expect(after.get(name)).toEqual(data);
  }
  const path = process.env.S03_INTERCHANGE_OUT;
  if (path) {
    await mkdir(path, { recursive: true });
    await writeFile(join(path, 'DesktopStyles.xlsx'), new Uint8Array(await output.blob.arrayBuffer()));
  }
});

it('rejects XML document types instead of resolving them', () => {
  expect(() => roundTripStyles('<!DOCTYPE styleSheet><styleSheet/>')).toThrow('document types');
});
