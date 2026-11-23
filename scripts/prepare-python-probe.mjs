import { execFile } from 'node:child_process';
import { createHash } from 'node:crypto';
import { copyFile, cp, mkdir, readFile, rename, rm, writeFile } from 'node:fs/promises';
import { resolve, sep } from 'node:path';
import { promisify } from 'node:util';
import { fileURLToPath } from 'node:url';
import { gunzipSync } from 'node:zlib';
import { TextWriter, Uint8ArrayReader, ZipReader } from '@zip.js/zip.js';

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

async function stagePyodideLicense() {
  const notice = lock.license;
  if (!notice || !/^https:\/\//.test(notice.source)) throw new Error('The pinned Pyodide license source is missing.');
  const path = new URL(notice.name, directory);
  if (!await matches(path, notice)) {
    const temporary = new URL(`${notice.name}.${process.pid}.part`, directory);
    try {
      await run('curl', ['--fail', '--silent', '--show-error', '--location', '--retry', '2',
        '--max-time', '120', notice.source, '--output', fileURLToPath(temporary)]);
      if (!await matches(temporary, notice)) throw new Error(`Checksum or size mismatch for ${notice.name}`);
      await rename(temporary, path);
    } finally {
      await rm(temporary, { force: true });
    }
  }
  const productionPath = new URL(notice.name, productionDirectory);
  if (!await matches(productionPath, notice)) {
    const temporary = new URL(`${notice.name}.${process.pid}.part`, productionDirectory);
    try {
      await copyFile(path, temporary);
      if (!await matches(temporary, notice)) throw new Error(`Production license copy checksum mismatch for ${notice.name}`);
      await rename(temporary, productionPath);
    } finally {
      await rm(temporary, { force: true });
    }
  }
}

async function generateNotices(runtimeDirectory) {
  const runtimePath = fileURLToPath(runtimeDirectory);
  const pyodideLock = JSON.parse(await readFile(new URL('pyodide-lock.json', runtimeDirectory), 'utf8'));
  const licenseDirectory = resolve(runtimePath, 'licenses');
  await rm(licenseDirectory, { recursive: true, force: true });
  await mkdir(licenseDirectory, { recursive: true });
  const rows = [];
  for (const [packageName, version] of Object.entries(lock.packages).sort(([left], [right]) => left.localeCompare(right))) {
    const packageMetadata = pyodideLock.packages[packageName];
    if (!packageMetadata || packageMetadata.version !== version || !packageMetadata.file_name?.endsWith('.whl')) {
      throw new Error(`The pinned package metadata is missing for ${packageName} ${version}.`);
    }
    const archivePath = new URL(packageMetadata.file_name, runtimeDirectory);
    const reader = new ZipReader(new Uint8ArrayReader(await readFile(archivePath)));
    try {
      const entries = await reader.getEntries();
      const metadataEntry = entries.find(entry => /\.dist-info\/METADATA$/i.test(entry.filename));
      if (!metadataEntry) throw new Error(`Wheel metadata is missing for ${packageName} ${version}.`);
      const distInfoPrefix = metadataEntry.filename.slice(0, -'METADATA'.length);
      const metadataText = await metadataEntry.getData(new TextWriter());
      const metadataLine = name => metadataText.match(new RegExp(`^${name}:\\s*(.+)$`, 'im'))?.[1]?.trim();
      const licenseFiles = entries.filter(entry => !entry.directory && entry.filename.startsWith(distInfoPrefix) &&
        /(?:license|licence|copying|notice|copyright)(?:[._-].*)?$/i.test(entry.filename.split('/').at(-1) ?? ''));
      const outputFiles = [];
      for (const entry of licenseFiles) {
        const relativePath = entry.filename.slice(distInfoPrefix.length).split('/');
        if (relativePath.some(part => !part || part === '.' || part === '..')) continue;
        const destination = resolve(licenseDirectory, packageName, ...relativePath);
        const packageRoot = resolve(licenseDirectory, packageName) + sep;
        if (!destination.startsWith(packageRoot)) throw new Error(`Unsafe license path in ${packageMetadata.file_name}.`);
        const contents = await entry.getData(new TextWriter());
        if (!contents.trim()) continue;
        await mkdir(resolve(destination, '..'), { recursive: true });
        await writeFile(destination, contents, 'utf8');
        outputFiles.push(`licenses/${packageName}/${relativePath.join('/')}`);
      }
      const license = metadataLine('License-Expression') ?? metadataLine('License') ??
        metadataText.match(/^Classifier:\s*License ::\s*(.+)$/im)?.[1]?.trim() ?? 'See the packaged wheel metadata';
      const sourceUrl = metadataText.match(/^Project-URL:\s*(?:source|repository|homepage|homepage|home)\s*,\s*(https?:\/\/\S+)$/im)?.[1];
      rows.push({ packageName, version, license, sourceUrl, outputFiles });
    } finally {
      await reader.close();
    }
  }

  async function fetchLocked(source, name, asset) {
    const sourceCache = new URL(`.notice-source/${name}`, runtimeDirectory);
    if (await matches(sourceCache, asset)) return await readFile(sourceCache);
    await rm(sourceCache, { force: true });
    await mkdir(new URL('.notice-source/', runtimeDirectory), { recursive: true });
    const temporary = new URL(`.notice-source/${name}.${process.pid}.part`, runtimeDirectory);
    try {
      await run('curl', ['--fail', '--silent', '--show-error', '--location', '--retry', '2',
        '--max-time', '120', source, '--output', fileURLToPath(temporary)]);
      if (!await matches(temporary, asset)) throw new Error(`Checksum or size mismatch for ${name}`);
      await rename(temporary, sourceCache);
      return await readFile(sourceCache);
    } finally {
      await rm(temporary, { force: true });
    }
  }

  function tarEntry(archive, wantedPath) {
    let offset = 0;
    while (offset + 512 <= archive.length) {
      const header = archive.subarray(offset, offset + 512);
      const name = header.subarray(0, 100).toString('utf8').replace(/\0.*$/, '');
      const prefix = header.subarray(345, 500).toString('utf8').replace(/\0.*$/, '');
      if (!name) break;
      const fullName = prefix ? `${prefix}/${name}` : name;
      const sizeText = header.subarray(124, 136).toString('ascii').replace(/\0.*$/, '').trim();
      const size = Number.parseInt(sizeText || '0', 8);
      const bodyStart = offset + 512;
      if (fullName === wantedPath || fullName.endsWith('/' + wantedPath)) {
        return archive.subarray(bodyStart, bodyStart + size);
      }
      offset = bodyStart + Math.ceil(size / 512) * 512;
    }
    throw new Error(`Supplemental package archive does not contain ${wantedPath}.`);
  }

  for (const supplemental of lock.supplementalNotices ?? []) {
    const sourceFiles = [];
    let sourceLicenseMissing = false;
    if (supplemental.archive) {
      const archive = supplemental.archive;
      const bytes = await fetchLocked(archive.source, archive.name, archive);
      try {
        const contents = tarEntry(gunzipSync(bytes), archive.licensePath).toString('utf8');
        sourceFiles.push({ output: archive.output, contents });
      } catch (error) {
        if (!archive.textMayBeAbsent) throw error;
        sourceLicenseMissing = true;
      }
    }
    for (const file of supplemental.files ?? []) {
      sourceFiles.push({ output: file.output, contents: (await fetchLocked(file.source, file.name, file)).toString('utf8') });
    }
    let row = rows.find(item => item.packageName === supplemental.package);
    if (!row) {
      row = { packageName: supplemental.package, version: supplemental.version, license: supplemental.license,
        sourceUrl: supplemental.sourceUrl, outputFiles: [] };
      rows.push(row);
    }
    if (sourceLicenseMissing) row.missingLicenseText = 'The checksum-pinned source archive also omits a license text.';
    for (const file of sourceFiles) {
      const outputParts = file.output.split('/');
      if (outputParts[0] !== 'licenses' || outputParts.some(part => !part || part === '.' || part === '..')) {
        throw new Error(`Unsafe supplemental license output path: ${file.output}`);
      }
      const destination = resolve(licenseDirectory, ...outputParts.slice(1));
      if (!destination.startsWith(licenseDirectory + sep)) throw new Error(`Unsafe supplemental license path: ${file.output}`);
      await mkdir(resolve(destination, '..'), { recursive: true });
      await writeFile(destination, file.contents, 'utf8');
      row.outputFiles.push(file.output);
    }
  }

  rows.unshift({ packageName: 'Pyodide', version: lock.version, license: 'MPL-2.0',
    sourceUrl: `https://github.com/pyodide/pyodide/tree/${lock.version}`, outputFiles: [lock.license.name] });
  const lines = [
    '# Bundled Python runtime notices',
    '',
    `Pyodide ${lock.version} and CPython ${lock.python} are bundled from the checksum-locked runtime. ` +
      `The Pyodide license is included as [PYODIDE-LICENSE.txt](./${lock.license.name}).`,
    '',
    'Each package is listed with its pinned version and the license declaration in its wheel metadata. ' +
      'Where the wheel includes license texts, the complete files are copied below `licenses/<package>/`. ' +
      'Some upstream wheels omit a text file; their declared license and project source are recorded here.',
    '',
    '| Package | Version | Declared license | Project source | Included license texts |',
    '| --- | --- | --- | --- | --- |',
    ...rows.map(row => `| ${row.packageName} | ${row.version} | ${row.license.replaceAll('|', '\\|')} | ` +
      `${row.sourceUrl ? `[upstream](${row.sourceUrl})` : 'Not listed in wheel metadata'} | ` +
      `${row.outputFiles.length ? row.outputFiles.map(path => `[${path}](./${path})`).join('<br>') : row.missingLicenseText ?? 'None in wheel'} |`),
    '',
    'This file is generated by `scripts/prepare-python-probe.mjs` from the exact locked wheel files. ' +
      'Review the upstream project license if a wheel omits the license text.',
    '',
  ];
  await writeFile(resolve(runtimePath, 'LICENSES.md'), lines.join('\n'), 'utf8');
}

await stagePyodideLicense();
await generateNotices(directory);
await rm(new URL('licenses/', productionDirectory), { recursive: true, force: true });
await cp(new URL('licenses/', directory), new URL('licenses/', productionDirectory), { recursive: true, force: true });
await copyFile(new URL('LICENSES.md', directory), new URL('LICENSES.md', productionDirectory));
console.log(`Verified and staged ${lock.assets.length} Pyodide ${lock.version} assets and generated package license notices.`);
