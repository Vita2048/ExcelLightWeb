import { LocalWorkspaceError, type WorkspaceFileHandle } from './workspace-store';

const nativeType = {
  description: 'Excel-Lite workbook',
  accept: { 'application/zip': ['.exlt'] },
};

interface PickerWindow extends Window {
  showOpenFilePicker?: (options?: {
    readonly multiple?: boolean;
    readonly types?: readonly [{ readonly description: string; readonly accept: { readonly [mime: string]: readonly string[] } }];
  }) => Promise<readonly WorkspaceFileHandle[]>;
  showSaveFilePicker?: (options?: {
    readonly suggestedName?: string;
    readonly types?: readonly [{ readonly description: string; readonly accept: { readonly [mime: string]: readonly string[] } }];
  }) => Promise<WorkspaceFileHandle>;
}

export interface OpenedNativeFile {
  readonly file: File;
  readonly handle?: WorkspaceFileHandle;
}

export interface ExternalFileVersion {
  readonly size: number;
  readonly lastModified: number;
}

export function hasNativeOpenPicker(): boolean {
  return typeof (window as PickerWindow).showOpenFilePicker === 'function';
}

export function hasNativeSavePicker(): boolean {
  return typeof (window as PickerWindow).showSaveFilePicker === 'function';
}

function isAbort(error: unknown): boolean {
  return error instanceof DOMException && error.name === 'AbortError';
}

function pickWithInput(): Promise<OpenedNativeFile | null> {
  return new Promise((resolve, reject) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.exlt,application/zip';
    input.setAttribute('aria-label', 'Open Excel-Lite workbook');
    Object.assign(input.style, { position: 'fixed', left: '-10000px', top: '0', width: '1px', height: '1px' });
    document.body.append(input);
    let settled = false;
    const finish = (result: OpenedNativeFile | null) => {
      if (settled) return;
      settled = true;
      input.remove();
      resolve(result);
    };
    input.addEventListener('change', () => {
      const file = input.files?.[0];
      finish(file ? { file } : null);
    }, { once: true });
    input.addEventListener('cancel', () => finish(null), { once: true });
    input.addEventListener('error', () => {
      if (settled) return;
      settled = true;
      input.remove();
      reject(new LocalWorkspaceError('The file picker could not read the selected file.'));
    }, { once: true });
    input.click();
  });
}

export async function pickNativeOpenFile(): Promise<OpenedNativeFile | null> {
  const picker = (window as PickerWindow).showOpenFilePicker;
  if (!picker) return pickWithInput();
  try {
    const [handle] = await picker.call(window, { multiple: false, types: [nativeType] });
    if (!handle) return null;
    return { file: await handle.getFile(), handle };
  } catch (error) {
    if (isAbort(error)) return null;
    throw error;
  }
}

export async function pickNativeSaveFile(suggestedName: string): Promise<WorkspaceFileHandle | null> {
  const picker = (window as PickerWindow).showSaveFilePicker;
  if (!picker) return null;
  try {
    return await picker.call(window, { suggestedName, types: [nativeType] });
  } catch (error) {
    if (isAbort(error)) return null;
    throw error;
  }
}

export function fileVersion(file: File): ExternalFileVersion {
  return { size: file.size, lastModified: file.lastModified };
}

export async function writeNativeFile(handle: WorkspaceFileHandle, packageBlob: Blob,
  expected?: ExternalFileVersion): Promise<ExternalFileVersion> {
  const permission = { mode: 'readwrite' as const };
  let state = await handle.queryPermission?.(permission);
  if (state !== 'granted') state = await handle.requestPermission?.(permission);
  if (state !== 'granted') throw new LocalWorkspaceError('File access was denied. The linked file was not changed.');

  if (expected) {
    const current = fileVersion(await handle.getFile());
    if (current.size !== expected.size || current.lastModified !== expected.lastModified) {
      throw new LocalWorkspaceError('The linked file changed outside ExcelLight. Your file was not overwritten. Use Save As to choose another file.');
    }
  }

  const writable = await handle.createWritable({ keepExistingData: false });
  try {
    await writable.write(packageBlob);
    await writable.close();
  } catch (error) {
    try { await writable.abort?.(error); } catch { /* Preserve the write failure. */ }
    throw error;
  }
  const written = await handle.getFile();
  if (written.size !== packageBlob.size) {
    throw new LocalWorkspaceError('The browser could not verify the saved file size. Keep the local snapshot or export a backup.');
  }
  return fileVersion(written);
}

export function downloadNativePackage(packageBlob: Blob, name: string): void {
  const normalized = name.toLowerCase().endsWith('.exlt') ? name : `${name}.exlt`;
  const url = URL.createObjectURL(packageBlob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = normalized;
  anchor.hidden = true;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 30_000);
}
