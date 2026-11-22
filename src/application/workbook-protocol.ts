import type { CellCache, CellReference, WorkbookDocument } from '../core/workbook';

export const workbookProtocolVersion = 1;

export interface SnapshotBasis {
  readonly sessionId: string;
  readonly revision: number;
  readonly generation: number;
  /** Opaque authored-state identity scoped to a session. */
  readonly sourceFingerprint: string;
}

export interface WorkbookSnapshot extends SnapshotBasis {
  readonly authoredRevision: number;
  readonly document: WorkbookDocument;
}

export interface WorkerBasis extends SnapshotBasis {
  readonly protocolVersion: typeof workbookProtocolVersion;
  readonly requestId: string;
  readonly runtimeFingerprint: string;
}

export interface NativeCalculationRequest extends WorkerBasis {
  readonly kind: 'calculate-native';
  readonly document: WorkbookDocument;
}

export interface CacheUpdate {
  readonly target: CellReference;
  /** Omit cache to remove a previously cached result. */
  readonly cache?: CellCache;
}

export interface NativeCalculationReply extends WorkerBasis {
  readonly kind: 'native-calculated';
  readonly updates: readonly CacheUpdate[];
}

export type NativeCalculationScope = 'automatic' | 'observe' | 'workbook' | { readonly sheetId: string } |
  { readonly targets: readonly string[] };

export interface NativeWorkerRequest extends NativeCalculationRequest {
  readonly scope: NativeCalculationScope;
  readonly acceptedRequestId?: string;
  readonly locale: string;
}

export interface NativeWorkerReply extends NativeCalculationReply {
  readonly staleCellIds: readonly string[];
}
