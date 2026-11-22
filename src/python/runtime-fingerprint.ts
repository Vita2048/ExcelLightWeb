import lock from './runtime-lock.json';

/** Keep persisted result freshness aligned with the pinned browser runtime. */
export const pythonRuntimeFingerprint = `pyodide-${lock.version}-python-${lock.python}-numpy-${lock.packages.numpy}` +
  `-pandas-${lock.packages.pandas}-matplotlib-${lock.packages.matplotlib}`;

