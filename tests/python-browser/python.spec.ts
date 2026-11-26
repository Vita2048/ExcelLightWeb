import { execFileSync } from 'node:child_process';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { expect, test, type Page } from '@playwright/test';
import lock from '../../src/python/runtime-lock.json' with { type: 'json' };
import type { RunOutput, Scalar, Snapshot } from '../../src/python/probe-protocol';
import salesExpected from '../fixtures/sales/Expected-values.json' with { type: 'json' };
import slaExpected from '../fixtures/sla/Expected-values.json' with { type: 'json' };

type Fixture = {
  scripts: Record<string, string>;
  snapshots: Record<string, Snapshot>;
  baselinePivots: Scalar[][][];
};

function fixture(name: string): Fixture {
  const reader = fileURLToPath(new URL('./read-fixtures.py', import.meta.url));
  return JSON.parse(execFileSync(process.env.PYTHON ?? 'python', ['-B', reader, name], { encoding: 'utf8' })) as Fixture;
}

function run(page: Page, source: string, snapshots: Record<string, Snapshot> = {}) {
  return page.evaluate(({ source, snapshots }) => window.pythonProbe.run(source, snapshots), { source, snapshots });
}

function table(output: RunOutput) {
  if (output.result.kind !== 'table') throw new Error('Expected a table');
  return output.result;
}

function withHeaders(rows: Scalar[][], width: number): Snapshot {
  const [head, ...data] = rows;
  if (!head) throw new Error('Expected headers in fixture result');
  return { headers: head.slice(0, width).map(String), rows: data.map(row => row.slice(0, width)) };
}

function png(output: RunOutput, dimensions: [number, number]) {
  if (output.result.kind !== 'png') throw new Error('Expected an Agg figure');
  const bytes = Buffer.from(output.result.base64, 'base64');
  expect(bytes.subarray(0, 8).toString('hex')).toBe('89504e470d0a1a0a');
  expect([bytes.readUInt32BE(16), bytes.readUInt32BE(20)]).toEqual(dimensions);
  return bytes;
}

async function start(page: Page) {
  const unexpected: string[] = [];
  const origin = 'http://127.0.0.1:' + (process.env.PLAYWRIGHT_TEST_PORT ?? '4173');
  await page.route(url => url.origin !== origin, route => {
    unexpected.push(route.request().url());
    return route.abort();
  });
  await page.goto('/ExcelLightWeb/');
  const start = performance.now();
  const versions = await page.evaluate(() => window.pythonProbe.start());
  await test.info().attach('runtime', { body: JSON.stringify({ versions, coldMs: performance.now() - start }), contentType: 'application/json' });
  expect(versions).toEqual({
    python: lock.python, numpy: lock.packages.numpy, pandas: lock.packages.pandas,
    matplotlib: lock.packages.matplotlib,
  });
  return unexpected;
}

function script(data: Fixture, key: string): string {
  const source = data.scripts[key];
  if (!source) throw new Error(`Missing script at ${key}`);
  return source;
}

test('executes all six unmodified sales scripts with the fixture snapshots', async ({ page }) => {
  const unexpected = await start(page);
  const sales = fixture('docs/showcase/Showcase.exlt');
  const locations = [
    ['Sales Analysis', 3, 1], ['Rankings', 3, 1], ['Rankings', 3, 5], ['Charts', 3, 1],
  ] as const;
  expect(Object.keys(sales.scripts)).toHaveLength(6);
  for (const [sheet, row, column] of locations) {
    const result = table(await run(page, script(sales, `${sheet}!${row}:${column}`), sales.snapshots));
    for (const [i, values] of result.rows.entries()) {
      for (const [j, value] of values.entries()) {
        const expected = salesExpected.cells.find(cell => cell.sheet === sheet &&
          cell.row === row + i + 1 && cell.column === column + j);
        expect(expected, `${sheet} row ${row + i + 1} column ${column + j}`).toBeDefined();
        if (typeof value === 'number') expect(value).toBeCloseTo(Number(expected!.value), 8);
        else expect(value).toBe(expected!.value);
      }
    }
    if (sheet === 'Charts') {
      sales.snapshots['Charts!A3:D8'] = { headers: result.headers!, rows: result.rows.slice(0, 5) };
    }
  }
  const monthly = await run(page, script(sales, 'Charts!3:5'), sales.snapshots);
  png(monthly, [960, 576]);
  const monthlyTotals = JSON.parse(monthly.stdout.trim().replace(/^Monthly revenue: /, '')) as Record<string, number>;
  expect(monthlyTotals).toEqual(Object.fromEntries(salesExpected.monthly.map(([month, value]) => [month, Number(value)])));
  png(await run(page, script(sales, 'Charts!22:5'), sales.snapshots), [960, 576]);
  expect(unexpected).toEqual([]);
});

test('executes the unchanged SLA analysis and charts for both scenarios', async ({ page }) => {
  const unexpected = await start(page);
  const sla = fixture('docs/showcase/Support-SLA.exlt');
  for (const [path, key] of [['analysis', 'Analysis!6:2'], ['trend', 'Analysis!52:2'], ['heatmap', 'Analysis!53:2']] as const) {
    const file = readFileSync(new URL(`../fixtures/sla/${path}.py`, import.meta.url), 'utf8');
    expect(script(sla, key).replaceAll('\r\n', '\n').trimEnd()).toBe(file.replaceAll('\r\n', '\n').trimEnd());
  }
  const targets = sla.snapshots['ResponseTargets[#Data]'];
  if (!targets || !('rows' in targets)) throw new Error('Missing SLA targets');
  const targetIndex = targets.headers.indexOf('Target hrs');
  const p2 = targets.rows.find(row => row[0] === 'P2');
  if (targetIndex < 0 || !p2) throw new Error('Missing P2 target');
  const charts: Buffer[][] = [];
  for (const scenario of slaExpected.scenarios) {
    p2[targetIndex] = scenario.p2TargetHours;
    const report = table(await run(page, script(sla, 'Analysis!6:2'), sla.snapshots)).rows;
    expect(report).toHaveLength(43);
    const kpis = [scenario.kpis.opened, scenario.kpis.breachRate, scenario.kpis.open,
      scenario.kpis.resolutionP90Hours, scenario.pending, scenario.assessed];
    kpis.forEach((value, i) => expect(report[i + 1]?.[1]).toBeCloseTo(value, 9));
    if (scenario.p2TargetHours === 4) {
      for (const [index, first] of [21, 30].entries()) {
        for (let row = 0; row < 5; row++) {
          for (let col = 0; col < 6; col++) {
            const wanted = sla.baselinePivots[index]![row]![col];
            if (typeof wanted === 'number') expect(report[first + row]?.[col + 1]).toBeCloseTo(wanted, 8);
            else expect(report[first + row]?.[col + 1]).toBe(wanted);
          }
        }
      }
    }
    sla.snapshots['Analysis!B16:E22'] = withHeaders(report.slice(10, 17), 4);
    sla.snapshots['Analysis!B26:H31'] = withHeaders(report.slice(20, 26), 7);
    charts.push([
      png(await run(page, script(sla, 'Analysis!52:2'), sla.snapshots), [777, 464]),
      png(await run(page, script(sla, 'Analysis!53:2'), sla.snapshots), [777, 464]),
    ]);
  }
  expect(charts[0]![0]!.equals(charts[1]![0]!)).toBe(false);
  expect(charts[0]![1]!.equals(charts[1]![1]!)).toBe(false);
  p2[targetIndex] = 0;
  await expect(run(page, script(sla, 'Analysis!6:2'), sla.snapshots)).rejects.toThrow(slaExpected.invalidTarget.errorContains);
  expect(unexpected).toEqual([]);
});

test('normalizes results and redraws a private serialized Figure', async ({ page }) => {
  const unexpected = await start(page);
  expect((await run(page, 'result = 1\nprint("output")\nNone')).result).toEqual({ kind: 'scalar', value: null });
  expect((await run(page, 'import numpy as np\nresult = np.int64(42)')).result).toEqual({ kind: 'scalar', value: 42 });
  expect(table(await run(page, 'import pandas as pd\npd.Series([1, pd.NA], name="Count")'))).toMatchObject({
    headers: ['Count'], rows: [[1], [null]],
  });
  expect(table(await run(page, 'import numpy as np\nnp.array([[1, 2], [3, 4]])')).rows).toEqual([[1, 2], [3, 4]]);
  await expect(run(page, 'float("inf")')).rejects.toThrow('Unsupported result value');
  await expect(run(page, 'result')).rejects.toThrow('NameError');
  png(await run(page, `
import pickle
from matplotlib import pyplot as plt
fig, ax = plt.subplots()
ax.plot([1, 2], [3, 4])
restored = pickle.loads(pickle.dumps(fig))
restored.set_dpi(100)
restored.set_size_inches(3.2, 2.4)
restored.tight_layout()
restored
`), [320, 240]);
  expect(unexpected).toEqual([]);
});

test('loads Arrow on demand and round-trips a nullable table through IPC', async ({ page }) => {
  const unexpected = await start(page);
  const started = performance.now();
  expect(await page.evaluate(() => window.pythonProbe.loadArrow())).toBe(lock.packages.pyarrow);
  await test.info().attach('arrow-startup', {
    body: JSON.stringify({ loadMs: performance.now() - started }),
    contentType: 'application/json',
  });
  expect((await run(page, `
import pyarrow as pa
import pyarrow.ipc as ipc
out = pa.BufferOutputStream()
table = pa.table({"count": [1, None, 3]})
with ipc.new_stream(out, table.schema) as writer:
    writer.write_table(table)
ipc.open_stream(out.getvalue()).read_all().equals(table)
`)).result).toEqual({ kind: 'scalar', value: true });
  expect(unexpected).toEqual([]);
});

test('reports cold runtime startup and warm sales-script runs', async ({ page }) => {
  const coldStart = performance.now();
  const unexpected = await start(page);
  const coldMs = performance.now() - coldStart;
  const sales = fixture('docs/showcase/Showcase.exlt');
  const source = script(sales, 'Sales Analysis!3:1');
  const warm: number[] = [];
  for (let index = 0; index < 20; index++) {
    const started = performance.now();
    const result = table(await run(page, source, sales.snapshots));
    warm.push(performance.now() - started);
    expect(result.rows.length).toBeGreaterThan(0);
  }
  const sorted = [...warm].sort((a, b) => a - b);
  const measurement = {
    runtime: lock,
    coldStartupMs: coldMs,
    warmRuns: { samples: warm.length, p50: sorted[9], p95: sorted[18], max: sorted.at(-1) },
  };
  await test.info().attach('python-timings', { body: JSON.stringify(measurement, null, 2), contentType: 'application/json' });
  mkdirSync('test-results/performance', { recursive: true });
  writeFileSync('test-results/performance/python-timings.json', JSON.stringify(measurement, null, 2) + '\n');
  expect(unexpected).toEqual([]);
});

test('terminates a CPU-bound cell at 30 seconds and restarts cleanly', async ({ page }) => {
  const unexpected = await start(page);
  const started = performance.now();
  await expect(run(page, 'while True:\n    pass')).rejects.toThrow('Python timed out after 30000 ms');
  expect(performance.now() - started).toBeGreaterThanOrEqual(29_900);
  await expect(run(page, '1')).rejects.toThrow('Python worker has stopped');
  await page.evaluate(() => window.pythonProbe.start());
  expect((await run(page, '2 + 3')).result).toEqual({ kind: 'scalar', value: 5 });
  expect(unexpected).toEqual([]);
});
