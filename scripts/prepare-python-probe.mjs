import { execFile } from 'node:child_process';
import { createHash } from 'node:crypto';
import { copyFile, mkdir, readFile, rename, rm } from 'node:fs/promises';
import { promisify } from 'node:util';
import { fileURLToPath } from 'node:url';

const run = promisify(execFile);
const lock = JSON.parse(await readFile(new URL('../src/python/runtime-lock.json', import.meta.url), 'utf8'));
const directory = new URL(`../tests/python-browser/public/runtime/pyodide/${lock.version}/`, import.meta.url);
const productionDirectory = new URL(`../public/runtime/pyodide/${lock.version}/`, import.meta.url);
await mkdir(directory, { recursive: true });
await mkdir(productionDirectory, { recursive: true });

async function matches(path, asset) {
  const bytes = await readFile(path).catch(error => {
    if (error.code === 'ENOENT') return null;
    throw error;
  });
  return bytes?.length === asset.bytes && createHash('sha256').update(bytes).digest('hex') === asset.sha256;
}

for (const asset of lock.assets) {
  const path = new URL(asset.name, directory);
  if (!await matches(path, asset)) {
    const temporary = new URL(`${asset.name}.${process.pid}.part`, directory);
    try {
      const url = new URL(asset.name, lock.source);
      await run('curl', ['--fail', '--silent', '--show-error', '--location', '--retry', '2',
        '--max-time', '120', url.href, '--output', fileURLToPath(temporary)]);
      if (!await matches(temporary, asset)) throw new Error(`Checksum or size mismatch for ${asset.name}`);
      await rename(temporary, path);
    } finally {
      await rm(temporary, { force: true });
    }
  }
  const productionPath = new URL(asset.name, productionDirectory);
  if (!await matches(productionPath, asset)) {
    const productionTemporary = new URL(`${asset.name}.${process.pid}.part`, productionDirectory);
    try {
      await copyFile(new URL(asset.name, directory), productionTemporary);
      if (!await matches(productionTemporary, asset)) throw new Error(`Production runtime copy checksum mismatch for ${asset.name}`);
      await rename(productionTemporary, productionPath);
    } finally {
      await rm(productionTemporary, { force: true });
    }
  }
}
console.log(`Verified and staged ${lock.assets.length} Pyodide ${lock.version} assets.`);
