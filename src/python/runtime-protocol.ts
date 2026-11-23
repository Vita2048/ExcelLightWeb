import type { CellCache, PythonFigure } from '../core/workbook';

export const pythonRuntimeProtocolVersion = 1;

export type PythonInputValue = string | number | boolean | null;
export interface PythonInputTable {
  readonly rows: readonly (readonly PythonInputValue[])[];
  /** Labels supplied from a workbook table, or from the first selected row for headers=True. */
  readonly headers?: readonly string[];
  /** A one cell range is scalar when headers=False. */
  readonly scalarCell?: boolean;
  readonly scalar?: PythonInputValue;
}
export interface PythonInputCall {
  readonly reference: string;
  readonly headers: boolean;
  readonly key: string;
}
export interface PythonInputBinding {
  readonly name: string;
  readonly key: string;
}

export interface PythonRuntimeInput {
  readonly calls: readonly PythonInputCall[];
  readonly bindings: readonly PythonInputBinding[];
  readonly values: Readonly<Record<string, PythonInputTable>>;
}

export type PythonRuntimeResult =
  | { readonly kind: 'scalar'; readonly value: PythonInputValue }
  | { readonly kind: 'table'; readonly headers: readonly string[] | null; readonly rows: readonly (readonly PythonInputValue[])[]; readonly notices: readonly string[] }
  | { readonly kind: 'figure'; readonly figure: PythonFigure };

export interface PythonRuntimeOutput {
  readonly status: 'succeeded' | 'failed';
  readonly result?: PythonRuntimeResult;
  readonly stdout: string;
  readonly stderr: string;
  readonly error?: string;
  readonly outputTruncated: boolean;
}

export interface PythonRuntimeProgress {
  readonly kind: 'progress';
  readonly protocolVersion: typeof pythonRuntimeProtocolVersion;
  readonly requestId: string;
  readonly message: string;
}

export type PythonRuntimeRequest =
  | { readonly kind: 'init'; readonly protocolVersion: typeof pythonRuntimeProtocolVersion; readonly requestId: string; readonly runtimeUrl: string }
  | { readonly kind: 'run'; readonly protocolVersion: typeof pythonRuntimeProtocolVersion; readonly requestId: string; readonly source: string; readonly input: PythonRuntimeInput }
  | { readonly kind: 'render-figure'; readonly protocolVersion: typeof pythonRuntimeProtocolVersion; readonly requestId: string;
      readonly figureId: string; readonly width: number; readonly height: number; readonly dpi: number };

export type PythonRuntimeReply =
  | { readonly kind: 'ready'; readonly protocolVersion: typeof pythonRuntimeProtocolVersion; readonly requestId: string; readonly runtimeFingerprint: string }
  | { readonly kind: 'result'; readonly protocolVersion: typeof pythonRuntimeProtocolVersion; readonly requestId: string; readonly output: PythonRuntimeOutput }
  | { readonly kind: 'error'; readonly protocolVersion: typeof pythonRuntimeProtocolVersion; readonly requestId: string; readonly message: string }
  | PythonRuntimeProgress;

export function pythonCache(value: PythonInputValue): CellCache | undefined {
  if (value === null) return undefined;
  if (typeof value === 'number') return { kind: 'number', value, serializedText: String(value) };
  if (typeof value === 'boolean') return { kind: 'boolean', value };
  return { kind: 'text', value };
}
