export enum SaveFault {
  None = 0,
  DuringWrite = 1,
  BeforeReplace = 2,
}

export interface PackageLimits {
  readonly maxPackageBytes?: number;
  readonly maxEntryBytes?: number;
  readonly maxExpandedBytes?: number;
  readonly maxEntries?: number;
}

export const defaultPackageLimits: Required<PackageLimits> = Object.freeze({
  maxPackageBytes: 128 * 1024 * 1024,
  maxEntryBytes: 32 * 1024 * 1024,
  maxExpandedBytes: 512 * 1024 * 1024,
  maxEntries: 10_000,
});

export interface SaveRequest {
  readonly fault?: SaveFault;
  readonly beforeReplace?: () => void;
  readonly limits?: PackageLimits;
}
