import adapter from '../python/runtime-adapter.py?raw';
import { pythonRuntimeFingerprint } from '../python/runtime-fingerprint';
import { pythonRuntimeProtocolVersion, type PythonRuntimeReply, type PythonRuntimeRequest } from '../python/runtime-protocol';

type Pyodide = {
  loadPackage(packages: string[]): Promise<unknown>;
  runPython(source: string): unknown;
  globals: { set(name: string, value: string | number): void; delete(name: string): void };
};
type Loader = { loadPyodide(options: { indexURL: string }): Promise<Pyodide> };

let runtime: Pyodide | undefined;
function send(reply: PythonRuntimeReply) {
  self.postMessage(reply);
}

self.addEventListener('message', async (event: MessageEvent<PythonRuntimeRequest>) => {
  const request = event.data;
  if (!request || request.protocolVersion !== pythonRuntimeProtocolVersion || typeof request.requestId !== 'string') return;
  try {
    if (request.kind === 'init') {
      const index = new URL(request.runtimeUrl);
      if (index.origin !== self.location.origin) throw new Error('Python runtime must be loaded from this application origin.');
      if (!index.pathname.endsWith('/')) index.pathname += '/';
      const loaderUrl = new URL('pyodide.mjs', index).href;
      const { loadPyodide } = await import(/* @vite-ignore */ loaderUrl) as Loader;
      send({ kind: 'progress', protocolVersion: pythonRuntimeProtocolVersion, requestId: request.requestId,
        message: 'Loading Python runtime…' });
      runtime = await loadPyodide({ indexURL: index.href });
      const packages = [
        { name: 'NumPy', package: 'numpy' },
        { name: 'pandas', package: 'pandas' },
        { name: 'Matplotlib', package: 'matplotlib' },
      ];
      for (const [index, item] of packages.entries()) {
        send({ kind: 'progress', protocolVersion: pythonRuntimeProtocolVersion, requestId: request.requestId,
          message: `Loading ${item.name} (${index + 1} of ${packages.length})…` });
        await runtime.loadPackage([item.package]);
        send({ kind: 'progress', protocolVersion: pythonRuntimeProtocolVersion, requestId: request.requestId,
          message: `Loaded ${item.name} (${index + 1} of ${packages.length}).` });
      }
      runtime.runPython(adapter);
      send({ kind: 'ready', protocolVersion: pythonRuntimeProtocolVersion, requestId: request.requestId,
        runtimeFingerprint: pythonRuntimeFingerprint });
      return;
    }

    if (!runtime) throw new Error('Python runtime is not initialized.');
    if (request.kind === 'render-figure') {
      runtime.globals.set('_excellite_figure_id', request.figureId);
      runtime.globals.set('_excellite_figure_width', request.width);
      runtime.globals.set('_excellite_figure_height', request.height);
      runtime.globals.set('_excellite_figure_dpi', request.dpi);
      try {
        const json = runtime.runPython('render_figure(_excellite_figure_id, _excellite_figure_width, _excellite_figure_height, _excellite_figure_dpi)') as string;
        const output = JSON.parse(json);
        send({ kind: 'result', protocolVersion: pythonRuntimeProtocolVersion, requestId: request.requestId, output });
      } finally {
        runtime.globals.delete('_excellite_figure_id');
        runtime.globals.delete('_excellite_figure_width');
        runtime.globals.delete('_excellite_figure_height');
        runtime.globals.delete('_excellite_figure_dpi');
      }
      return;
    }

    runtime.globals.set('_excellite_source', request.source);
    runtime.globals.set('_excellite_input', JSON.stringify(request.input));
    try {
      const json = runtime.runPython('run(_excellite_source, _excellite_input)') as string;
      const output = JSON.parse(json);
      send({ kind: 'result', protocolVersion: pythonRuntimeProtocolVersion, requestId: request.requestId, output });
    } finally {
      runtime.globals.delete('_excellite_source');
      runtime.globals.delete('_excellite_input');
    }
  } catch (error) {
    send({ kind: 'error', protocolVersion: pythonRuntimeProtocolVersion, requestId: request.requestId,
      message: error instanceof Error ? error.message : String(error) });
  }
});
