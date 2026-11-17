/** Relative assets support both a repository Pages URL and a custom-domain root. */
export function deploymentBase(value = './'): string {
  if (value === './') return value;
  if (!/^\/(?:[A-Za-z0-9_-]+\/)*$/.test(value)) {
    throw new Error('VITE_BASE_PATH must be ./ or a slash-delimited path, e.g. /ExcelLightWeb/.');
  }
  return value;
}
