// Loaded only when the Python panel opens. The editor, Python language grammar, CSS and
// worker are bundled into this deployment; no CDN or remote asset is requested.
import * as monaco from 'monaco-editor/editor/editor.api';
import 'monaco-editor/languages/definitions/python/register.js';
import EditorWorker from 'monaco-editor/editor/editor.worker?worker';

type MonacoHost = typeof globalThis & {
  MonacoEnvironment?: { getWorker: () => Worker };
};

(globalThis as MonacoHost).MonacoEnvironment = { getWorker: () => new EditorWorker() };

export { monaco };
