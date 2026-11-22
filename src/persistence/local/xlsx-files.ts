import { LocalWorkspaceError, type WorkspaceFileHandle } from './workspace-store';

const xlsxType = {
  description: 'Excel workbook',
  accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] },
};

interface PickerWindow extends Window {
  showSaveFilePicker?: (options?: {
    readonly suggestedName?: string;
    readonly types?: readonly { readonly description: string; readonly accept: Readonly<Record<string, readonly string[]>> }[];
  }) => Promise<WorkspaceFileHandle>;
}

function isAbort(error: unknown): boolean {
  return error instanceof DOMException && error.name === 'AbortError';
}

export function hasXlsxSavePicker(): boolean {
  return typeof (window as PickerWindow).showSaveFilePicker === 'function';
}

/** A separate destination picker for XLSX export; never used by the native Save lifecycle. */
export async function pickXlsxSaveFile(suggestedName: string): Promise<WorkspaceFileHandle | null> {
  const picker = (window as PickerWindow).showSaveFilePicker;
  if (!picker) return null;
  try {
    return await picker.call(window, { suggestedName, types: [xlsxType] });
  } catch (error) {
    if (isAbort(error)) return null;
    throw error;
  }
}

export function downloadXlsxFile(blob: Blob, name: string): void {
  const normalized = name.toLowerCase().endsWith('.xlsx') ? name : `${name}.xlsx`;
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = normalized;
  anchor.hidden = true;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 30_000);
}

/** File System Access identity check; unavailable for input-element imports. */
export async function sameFileEntry(left: WorkspaceFileHandle, right: WorkspaceFileHandle): Promise<boolean> {
  const check = (left as WorkspaceFileHandle & { isSameEntry?: (other: unknown) => Promise<boolean> }).isSameEntry;
  if (typeof check !== 'function') return false;
  try {
    return await check.call(left, right);
  } catch (error) {
    if (error instanceof LocalWorkspaceError) throw error;
    return false;
  }
}
