# ExcelLight Web

Browser-only port of the ExcelLight desktop workbook application: React and Canvas in a dark-only shell, with workbook model, calculation, file adapters and Python all running in the browser. It ships as static files on GitHub Pages with no application backend, telemetry or cloud workbook storage.

Live site: <https://vita2048.github.io/ExcelLightWeb/>

Implemented behavior includes the File backstage (New, Open, Save, Save As, Export, Close, Recent and local recovery), worksheet editing with clipboard, Fill, Clear and AutoSum, sheet and axis structure edits, the 31-function native calculator with Automatic/Manual modes, number formats, fonts, fills, borders, alignment and merges, workbook names and tables, sort and filter with SUBTOTAL, Find/Replace, zoom and freeze panes, native `.exlt` and reviewed XLSX import/export, browser-local Python through Pyodide with `=PY` anchors, spills and Matplotlib figure placements, and offline runtime caching with a user-controlled update flow.

## Documentation

- [Architecture](docs/Architecture.md): components, contracts and browser adaptations.
- [Features](docs/Features.md): the implemented behavior inventory with test evidence.
- [Development](docs/Development.md): commands, test layers and performance budgets.
- [Deployment](docs/Deployment.md): GitHub Pages publishing and runtime packaging.
- [Known issues](docs/Known-Issues.md): open limitations and unverified checks.
- [Runtime dependencies](docs/dependencies.md): pinned Python packages and notices.
- [ADR-003](docs/decisions/ADR-003-formula-and-file-adapters.md): formula and file adapter decisions.
- [Agent instructions](AGENTS.md).

The two sample workbooks are in [docs/showcase](docs/showcase). Only `Support-SLA.exlt` is bundled in the production build.

## Local commands

Use Node 22.12+ in the Node 22 release line and npm. The lockfile pins exact dependency resolutions.

```powershell
npm ci
npm run dev
npm run check
npx playwright install chromium firefox webkit
npm run test:e2e
npm run test:worksheet
```

`npm run check` runs the model and app typechecks, the Vitest suites and the production build. `test:e2e` serves the existing `dist`, so run `check` or `build` first. `npm run preview` serves the production build for manual inspection. No backend, database service, API keys or cloud account is needed for local development.
