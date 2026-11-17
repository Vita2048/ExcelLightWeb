import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

import manifest from '../fixtures/manifest.json';
import sales from '../fixtures/sales/Expected-values.json';
import sla from '../fixtures/sla/Expected-values.json';
import formula from '../fixtures/formulas/Expected-values.json';
import source from '../../docs/source-baseline/source-manifest.json';

const fixtureRoot = new URL('../fixtures/', import.meta.url);
const baselineRoot = new URL('../../docs/source-baseline/', import.meta.url);

function sha256(bytes: Buffer): string {
  return createHash('sha256').update(bytes).digest('hex');
}

describe('desktop baseline assets', () => {
  it('preserves exact fixture bytes and provenance across checkouts', () => {
    expect(new Set(manifest.files.map(file => file.path)).size).toBe(manifest.files.length);
    expect(manifest.files.length).toBeGreaterThan(0);
    for (const file of manifest.files) {
      const bytes = readFileSync(new URL(file.path, fixtureRoot));
      expect(bytes.length, file.path).toBe(file.bytes);
      expect(sha256(bytes), file.path).toBe(file.sha256);
      expect(file.sourcePath, file.path).toBeTruthy();
      expect(file.expectedValues, file.path).toBeTruthy();
    }
  });

  it('preserves frozen desktop source snapshots', () => {
    const snapshots = source.files.filter(file => file.snapshot !== undefined);
    expect(snapshots.length).toBeGreaterThan(0);
    for (const file of snapshots) {
      const bytes = readFileSync(new URL(file.snapshot!, baselineRoot));
      expect(sha256(bytes), file.path).toBe(file.sha256);
    }
  });

  it('loads independent sales expectations tied to the source XLSX', () => {
    expect(sales.source_sha256).toBe(sha256(readFileSync(new URL('sales/showcase_data.xlsx', fixtureRoot))));
    const coordinates = sales.cells.map(({ sheet, row, column }) => JSON.stringify([sheet, row, column]));
    expect(coordinates).toHaveLength(342);
    expect(new Set(coordinates).size).toBe(coordinates.length);
    expect(sales.monthly).toHaveLength(3);
  });

  it('loads SLA inputs, exact baseline KPIs and the P2 edit expectations', () => {
    expect(sla.workbookSha256).toBe(sha256(readFileSync(new URL('sla/Support-SLA.exlt', fixtureRoot))));
    expect(sla.scenarios.map(scenario => scenario.p2TargetHours)).toEqual([4, 2]);
    expect(sla.scenarios.map(scenario => scenario.kpis.opened)).toEqual([610, 610]);
    expect(sla.scenarios.map(scenario => scenario.kpis.open)).toEqual([8, 8]);
    expect(sla.scenarios.map(scenario => +(scenario.kpis.breachRate * 100).toFixed(1))).toEqual([24.8, 32.1]);
    for (const scenario of sla.scenarios) {
      expect(scenario.assessed + scenario.pending).toBe(scenario.kpis.opened);
    }
  });

  it('loads the formula and reference cases for the evaluator gate', () => {
    expect(formula.cases.find(test => test.id === 'lazy-if')).toMatchObject({ value: 42 });
    expect(formula.cases.find(test => test.id === 'malformed-source')).toMatchObject({ retainedSource: '=SUM(' });
    expect(formula.cases.find(test => test.id === 'relative-copy')).toMatchObject({ translated: '=$B2+C$2+$B$1' });
  });
});
