import type { PythonFigure, WorkbookDocument } from '../../core/workbook';
import { PackageException } from './package-exception';

const pngSignature = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
const crcTable = new Uint32Array(256);
for (let n = 0; n < 256; n++) {
  let value = n;
  for (let bit = 0; bit < 8; bit++) value = (value & 1) !== 0 ? (value >>> 1) ^ 0xedb88320 : value >>> 1;
  crcTable[n] = value >>> 0;
}

function crc32(bytes: Uint8Array): number {
  let value = 0xffffffff;
  for (const byte of bytes) value = crcTable[(value ^ byte) & 0xff]! ^ (value >>> 8);
  return (value ^ 0xffffffff) >>> 0;
}

function readUint32(bytes: Uint8Array, offset: number): number {
  return ((bytes[offset]! * 0x1000000) + (bytes[offset + 1]! << 16) + (bytes[offset + 2]! << 8) + bytes[offset + 3]!) >>> 0;
}

function decodeBase64(value: string): Uint8Array {
  if (value.length > 22_369_624) throw new PackageException('Figure PNG exceeds 16 MiB.');
  try {
    const binary = globalThis.atob(value);
    if (binary.length > 16 * 1024 * 1024) throw new PackageException('Invalid/bounded figure PNG.');
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return bytes;
  } catch (error) {
    if (error instanceof PackageException) throw error;
    throw new PackageException('Invalid figure PNG encoding.');
  }
}

async function validatePng(figure: PythonFigure): Promise<void> {
  const bytes = decodeBase64(figure.pngBase64);
  if (bytes.length < 45 || !pngSignature.every((byte, index) => bytes[index] === byte)) {
    throw new PackageException('Invalid/bounded figure PNG.');
  }

  let offset = pngSignature.length;
  let header = false;
  let end = false;
  let channels = 0;
  const idat: Uint8Array[] = [];
  let idatLength = 0;

  while (offset < bytes.length) {
    if (bytes.length - offset < 12) throw new PackageException('Truncated PNG chunk.');
    const length = readUint32(bytes, offset);
    if (length > bytes.length - offset - 12) throw new PackageException('Invalid PNG chunk length.');
    const type = String.fromCharCode(bytes[offset + 4]!, bytes[offset + 5]!, bytes[offset + 6]!, bytes[offset + 7]!);
    const crcStart = offset + 4;
    const crcEnd = offset + 8 + length;
    if (crc32(bytes.subarray(crcStart, crcEnd)) !== readUint32(bytes, crcEnd)) {
      throw new PackageException('PNG checksum mismatch.');
    }

    const data = bytes.subarray(offset + 8, offset + 8 + length);
    if (!header && type !== 'IHDR') throw new PackageException('PNG header is missing.');
    if (type === 'IHDR') {
      if (header || length !== 13 || readUint32(data, 0) !== figure.width || readUint32(data, 4) !== figure.height ||
        data[8] !== 8 || (data[9] !== 2 && data[9] !== 6) || data[10] !== 0 || data[11] !== 0 || data[12] !== 0) {
        throw new PackageException('Unsupported PNG header; expected bounded RGB/RGBA, 8-bit, noninterlaced.');
      }
      channels = data[9] === 6 ? 4 : 3;
      header = true;
    } else if (type === 'IDAT') {
      if (length > 16 * 1024 * 1024 - idatLength) throw new PackageException('Invalid/bounded figure PNG.');
      idat.push(data.slice());
      idatLength += length;
    } else if (type === 'IEND') {
      if (length !== 0 || offset + 12 !== bytes.length) throw new PackageException('Invalid PNG end.');
      end = true;
    } else if (type !== 'tEXt' && type !== 'pHYs') {
      throw new PackageException('Unsupported PNG chunk; only static Matplotlib PNGs are accepted.');
    }

    offset += length + 12;
  }

  if (!header || !end || idatLength === 0) throw new PackageException('Incomplete PNG.');

  const compressed = new Uint8Array(idatLength);
  let compressedOffset = 0;
  for (const chunk of idat) {
    compressed.set(chunk, compressedOffset);
    compressedOffset += chunk.length;
  }

  const rowBytes = figure.width * channels;
  const rowStride = rowBytes + 1;
  const expectedLength = rowStride * figure.height;
  const reader = new Blob([compressed]).stream().pipeThrough(new DecompressionStream('deflate')).getReader();
  let total = 0;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      for (let i = 0; i < value.length; i++) {
        const position = total + i;
        if (position % rowStride === 0 && value[i]! > 4) throw new PackageException('Invalid PNG scanline.');
      }
      total += value.length;
      if (total > expectedLength) throw new PackageException('PNG decoded data exceeds dimensions.');
    }
  } catch (error) {
    await reader.cancel().catch(() => {});
    if (error instanceof PackageException) throw error;
    throw new PackageException('Invalid PNG compressed pixels.');
  }
  if (total !== expectedLength) throw new PackageException('Invalid PNG compressed pixels.');
}

export async function validateNativeFigurePayloads(document: WorkbookDocument): Promise<void> {
  for (const sheet of document.sheets) {
    for (const cell of sheet.cells) {
      if (cell.content.kind !== 'python' || !cell.content.block) continue;
      const block = cell.content.block;
      if (block.lastFigure) await validatePng(block.lastFigure);
      for (const placement of block.placements ?? []) {
        if (placement.renderedFigure) await validatePng(placement.renderedFigure);
      }
    }
  }
}
