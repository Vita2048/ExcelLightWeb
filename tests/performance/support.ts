import { mkdirSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import { fileURLToPath } from 'node:url';

export interface Measurement { readonly name: string; readonly samples: readonly number[] }

export function percentile(samples: readonly number[], fraction: number): number {
  if (samples.length === 0) return 0;
  const sorted = [...samples].sort((a, b) => a - b);
  return sorted[Math.min(sorted.length - 1, Math.max(0, Math.ceil(fraction * sorted.length) - 1))]!;
}

export function summarize(measurements: readonly Measurement[]) {
  return measurements.map(({ name, samples }) => ({
    name,
    samples: samples.length,
    min: round(Math.min(...samples)),
    p50: round(percentile(samples, 0.5)),
    p95: round(percentile(samples, 0.95)),
    max: round(Math.max(...samples)),
  }));
}

function round(value: number) {
  return Math.round(value * 100) / 100;
}

export function environment() {
  const cpu = os.cpus()[0];
  return {
    captured: new Date().toISOString(),
    platform: os.platform(),
    release: os.release(),
    arch: os.arch(),
    cpu: cpu?.model.trim(),
    cores: os.cpus().length,
    memoryGiB: round(os.totalmem() / 1024 ** 3),
    node: process.version,
    vitest: process.env.npm_package_devDependencies_vitest ?? 'see package.json',
  };
}

/** Writes the measurement next to Playwright output; test-results/ is ignored by Git. */
export function record(name: string, payload: unknown) {
  const directory = fileURLToPath(new URL('../../test-results/performance/', import.meta.url));
  mkdirSync(directory, { recursive: true });
  const file = directory + name + '.json';
  writeFileSync(file, JSON.stringify(payload, null, 2) + '\n');
  return file;
}
