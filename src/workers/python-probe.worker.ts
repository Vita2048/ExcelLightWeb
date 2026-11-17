import adapter from '../python/probe.py?raw';
import type { ProbeRequest, ProbeResponse, RunOutput, Versions } from '../python/probe-protocol';

type Pyodide = {
  loadPackage(packages: string[]): Promise<unknown>;
  runPython(source: string): unknown;
  globals: { set(name: string, value: string): void; delete(name: string): void };
};
type Loader = { loadPyodide(options: { indexURL: string }): Promise<Pyodide> };
let runtime: Pyodide | undefined;

function send(response: ProbeResponse) {
  self.postMessage(response);
}

self.addEventListener('message', async (event: MessageEvent<ProbeRequest>) => {
  const request = event.data;
  try {
    if (request.kind === 'init') {
      const index = new URL(request.runtimeUrl);
      if (index.origin !== self.location.origin) throw new Error('Python runtime must be same-origin');
      const loaderUrl = new URL('pyodide.mjs', index).href;
      const { loadPyodide } = await import(/* @vite-ignore */ loaderUrl) as Loader;
      runtime = await loadPyodide({ indexURL: index.href });
      await runtime.loadPackage(['numpy', 'pandas', 'matplotlib']);
      runtime.runPython(adapter);
      const versions = JSON.parse(runtime.runPython(`
import platform
json.dumps({
    "python": platform.python_version(), "numpy": np.__version__,
    "pandas": pd.__version__, "matplotlib": matplotlib.__version__,
})
`) as string) as Versions;
      send({ kind: 'ready', versions });
    } else if (request.kind === 'load-arrow') {
      if (!runtime) throw new Error('Python worker is not initialized');
      await runtime.loadPackage(['pyarrow']);
      const version = runtime.runPython('import pyarrow as _probe_arrow; _probe_arrow.__version__') as string;
      send({ kind: 'arrow-ready', version });
    } else {
      if (!runtime) throw new Error('Python worker is not initialized');
      runtime.globals.set('_probe_source', request.source);
      runtime.globals.set('_probe_snapshots', JSON.stringify(request.snapshots));
      try {
        const json = runtime.runPython('run(_probe_source, _probe_snapshots)') as string;
        send({ kind: 'result', output: JSON.parse(json) as RunOutput });
      } finally {
        runtime.globals.delete('_probe_source');
        runtime.globals.delete('_probe_snapshots');
      }
    }
  } catch (error) {
    send({ kind: 'error', message: String(error) });
  }
});
