import { createHash } from 'node:crypto';
import { readdir, readFile, writeFile } from 'node:fs/promises';
import { resolve, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { gzipSync } from 'node:zlib';

const root = resolve(fileURLToPath(new URL('../dist/', import.meta.url)));
const runtimeRoot = `runtime/pyodide/`;
const workerTemplate = await readFile(new URL('./service-worker.template.js', import.meta.url), 'utf8');

async function listFiles(directory) {
  const result = [];
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    const absolute = resolve(directory, entry.name);
    if (entry.isDirectory()) result.push(...await listFiles(absolute));
    else if (entry.isFile()) result.push(absolute);
  }
  return result;
}

function sha256(bytes) {
  return createHash('sha256').update(bytes).digest('hex');
}

const files = (await listFiles(root)).sort();
const shell = [];
const runtime = [];
let totalBytes = 0;
for (const absolute of files) {
  const path = relative(root, absolute).split(sep).join('/');
  const bytes = await readFile(absolute);
  totalBytes += bytes.byteLength;
  const item = { url: path, bytes: bytes.byteLength, sha256: sha256(bytes) };
  if (path.startsWith(runtimeRoot)) runtime.push(item);
  else if (path !== 'service-worker.js' && path !== 'offline-manifest.json') shell.push(item);
}
if (!shell.some(asset => asset.url === 'index.html')) throw new Error('Production build is missing index.html.');
if (!runtime.length) throw new Error('Production build is missing the locked Python runtime.');

const buildId = sha256(Buffer.from(JSON.stringify(shell.map(({ url, sha256: hash }) => [url, hash]))));
const runtimeFingerprint = sha256(Buffer.from(JSON.stringify(runtime.map(({ url, sha256: hash }) => [url, hash]))));
const manifest = { schemaVersion: 1, buildId, runtimeFingerprint, shell, runtime };
await writeFile(resolve(root, 'offline-manifest.json'), JSON.stringify(manifest, null, 2) + '\n');
if (!workerTemplate.includes('__BUILD_ID__')) throw new Error('Service worker template has no build ID marker.');
await writeFile(resolve(root, 'service-worker.js'), workerTemplate.replaceAll('__BUILD_ID__', buildId));

const artifactFiles = await listFiles(root);
totalBytes = 0;
let totalGzipBytes = 0;
for (const absolute of artifactFiles) {
  const bytes = await readFile(absolute);
  totalBytes += bytes.byteLength;
  totalGzipBytes += gzipSync(bytes, { level: 9 }).byteLength;
}
const format = bytes => `${(bytes / 1024 / 1024).toFixed(2)} MiB`;
console.log(`Offline manifest: ${shell.length} app assets (${format(shell.reduce((sum, asset) => sum + asset.bytes, 0))}); ` +
  `${runtime.length} runtime/license assets (${format(runtime.reduce((sum, asset) => sum + asset.bytes, 0))}).`);
console.log(`Complete Pages artifact: ${format(totalBytes)} uncompressed; ${format(totalGzipBytes)} gzip across ${artifactFiles.length} files.`);
