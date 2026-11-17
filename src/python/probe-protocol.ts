export type Scalar = string | number | boolean | null;
export type Snapshot = { value: Scalar } | { headers: string[]; rows: Scalar[][] };
export type Result =
  | { kind: 'scalar'; value: Scalar }
  | { kind: 'table'; headers: string[] | null; rows: Scalar[][] }
  | { kind: 'png'; base64: string };
export type RunOutput = { result: Result; stdout: string; stderr: string };

export type ProbeRequest =
  | { kind: 'init'; runtimeUrl: string }
  | { kind: 'load-arrow' }
  | { kind: 'run'; source: string; snapshots: Record<string, Snapshot> };
export type Versions = { python: string; numpy: string; pandas: string; matplotlib: string };
export type ProbeResponse =
  | { kind: 'ready'; versions: Versions }
  | { kind: 'arrow-ready'; version: string }
  | { kind: 'result'; output: RunOutput }
  | { kind: 'error'; message: string };
