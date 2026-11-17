import { BlobReader, BlobWriter, Uint8ArrayReader, ZipReader, ZipWriter } from '@zip.js/zip.js';

const MiB = 1024 * 1024;
export type ZipBudget = {
  compressedBytes: number;
  expandedBytes: number;
  entryBytes: number;
  entries: number;
};
export const desktopZipBudget: ZipBudget = {
  compressedBytes: 128 * MiB, expandedBytes: 512 * MiB, entryBytes: 32 * MiB, entries: 10_000,
};
/** This probe retains all decoded parts for comparison. Production should visit sequentially. */
export const probeZipBudget: ZipBudget = {
  compressedBytes: 16 * MiB, expandedBytes: 32 * MiB, entryBytes: 4 * MiB, entries: 512,
};

export class ExpansionBudget {
  private total = 0;
  constructor(private readonly limits: ZipBudget) {}

  begin(name: string, reported: number) {
    let size = 0;
    const check = (bytes: number) => {
      if (!Number.isSafeInteger(bytes) || bytes < 0 || size + bytes > this.limits.entryBytes ||
          this.total + bytes > this.limits.expandedBytes) {
        throw new Error(`ZIP expanded byte limit exceeded for ${name}`);
      }
    };
    check(reported);
    return (bytes: number) => {
      check(bytes);
      size += bytes;
      this.total += bytes;
    };
  }
}

function normalizedPath(name: string) {
  const path = name.replaceAll('\\', '/');
  if (path.includes(':') || path.includes('\0') ||
      path.split('/').some(segment => !segment || segment === '.' || segment === '..')) {
    throw new Error(`Unsafe ZIP entry path: ${name}`);
  }
  return path;
}

/** One entry is active at a time. The byte budget is applied before forwarding each chunk. */
export async function visitZip(
  blob: Blob,
  limits: ZipBudget,
  openPart: (name: string) => { write(chunk: Uint8Array): void; finish(): void },
): Promise<void> {
  if (blob.size > limits.compressedBytes) throw new Error('ZIP compressed byte limit exceeded');
  const reader = new ZipReader(new BlobReader(blob), {
    strictness: 'strict', checkCrc32: true, checkOverlappingEntry: true,
    useWebWorkers: false,
  });
  const budget = new ExpansionBudget(limits);
  const names = new Set<string>();
  try {
    for await (const entry of reader.getEntriesGenerator()) {
      if (names.size >= limits.entries) throw new Error('ZIP entry count limit exceeded');
      if (entry.directory || entry.encrypted) throw new Error('ZIP directory records and encryption are outside the probe');
      const name = normalizedPath(entry.filename);
      const key = name.toLowerCase();
      if (names.has(key)) throw new Error(`Duplicate ZIP entry: ${name}`);
      names.add(key);
      const count = budget.begin(name, entry.uncompressedSize);
      const part = openPart(name);
      await entry.getData(new WritableStream<Uint8Array>({
        write(chunk) {
          count(chunk.byteLength);
          part.write(chunk);
        },
      }));
      part.finish();
    }
  } finally {
    await reader.close();
  }
}

export async function readProbeParts(blob: Blob, validatePath?: (name: string) => void) {
  const parts = new Map<string, Uint8Array<ArrayBuffer>>();
  await visitZip(blob, probeZipBudget, name => {
    validatePath?.(name);
    const chunks: Uint8Array[] = [];
    let size = 0;
    return {
      write(chunk) {
        chunks.push(chunk);
        size += chunk.byteLength;
      },
      finish() {
        const bytes = new Uint8Array(size);
        let offset = 0;
        for (const chunk of chunks) {
          bytes.set(chunk, offset);
          offset += chunk.byteLength;
        }
        parts.set(name, bytes);
      },
    };
  });
  return parts;
}

export async function writeProbeParts(parts: ReadonlyMap<string, Uint8Array<ArrayBuffer>>) {
  const writer = new ZipWriter(new BlobWriter('application/zip'), { useWebWorkers: false, zip64: false });
  for (const [name, data] of parts) await writer.add(name, new Uint8ArrayReader(data));
  return writer.close();
}
