import { afterEach, describe, expect, it, vi } from 'vitest';
import { LocalWorkspaceError, type WorkspaceFileHandle, type WorkspaceWritableFile } from '../../src/persistence/local';
import { pickNativeOpenFile, pickNativeSaveFile, writeNativeFile } from '../../src/persistence/local';

afterEach(() => vi.unstubAllGlobals());

function makeHandle(options: {
  permission?: PermissionState;
  requestPermission?: PermissionState;
  file?: File;
  write?: (blob: Blob) => void | Promise<void>;
  close?: () => void | Promise<void>;
}) {
  const calls = { getFile: 0, createWritable: 0, abort: 0, close: 0 };
  let file = options.file ?? new Blob(['old']) as File;
  const writable: WorkspaceWritableFile = {
    async write(blob) { await options.write?.(blob); file = blob as File; },
    async close() { calls.close++; await options.close?.(); },
    async abort() { calls.abort++; },
  };
  const handle: WorkspaceFileHandle = {
    name: 'book.exlt',
    async getFile() { calls.getFile++; return file; },
    async createWritable() { calls.createWritable++; return writable; },
    async queryPermission() { return options.permission ?? 'granted'; },
    async requestPermission() { return options.requestPermission ?? 'denied'; },
  };
  return { handle, calls };
}

describe('linked native file writes', () => {
  it('treats the native Open and Save As picker AbortError as cancellation', async () => {
    const openPicker = vi.fn().mockRejectedValue(new DOMException('Picker closed.', 'AbortError'));
    vi.stubGlobal('window', { showOpenFilePicker: openPicker });
    await expect(pickNativeOpenFile()).resolves.toBeNull();
    expect(openPicker).toHaveBeenCalledOnce();

    const savePicker = vi.fn().mockRejectedValue(new DOMException('Picker closed.', 'AbortError'));
    vi.stubGlobal('window', { showSaveFilePicker: savePicker });
    await expect(pickNativeSaveFile('Book1.exlt')).resolves.toBeNull();
    expect(savePicker).toHaveBeenCalledOnce();
  });

  it('does not create a writer when permission is denied or revoked', async () => {
    const { handle, calls } = makeHandle({ permission: 'prompt', requestPermission: 'denied' });
    await expect(writeNativeFile(handle, new Blob(['new']))).rejects.toThrow('File access was denied');
    expect(calls.createWritable).toBe(0);
  });

  it('refuses to overwrite a linked file whose size or modified time changed', async () => {
    const { handle, calls } = makeHandle({ file: new Blob(['external edit']) as File });
    await expect(writeNativeFile(handle, new Blob(['new']), { size: 3, lastModified: 100 })).rejects
      .toThrow('changed outside ExcelLight');
    expect(calls.createWritable).toBe(0);
  });

  it('closes the selected writer and verifies the saved byte length', async () => {
    const { handle, calls } = makeHandle({
      write: async blob => { expect(blob.size).toBe(3); },
      close: vi.fn(),
    });
    const version = await writeNativeFile(handle, new Blob(['new']));
    expect(version.size).toBe(3);
    expect(calls.createWritable).toBe(1);
    expect(calls.close).toBe(1);
  });

  it('aborts a partial writer if writing fails', async () => {
    const { handle, calls } = makeHandle({ write: () => { throw new Error('disk full'); } });
    await expect(writeNativeFile(handle, new Blob(['new']))).rejects.toThrow('disk full');
    expect(calls.abort).toBe(1);
  });
});
