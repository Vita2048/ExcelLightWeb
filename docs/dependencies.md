# Runtime dependency notices

The static app has no application backend. Browser computation uses the exact package versions in [`src/python/runtime-lock.json`](../src/python/runtime-lock.json); JavaScript package versions are pinned by [`package.json`](../package.json) and [`package-lock.json`](../package-lock.json).

`npm run prepare:python-runtime` downloads only files listed in the runtime lock, verifies each asset's byte count and SHA-256, verifies the pinned Pyodide license source, and extracts wheel license texts from those same verified archives. It emits the production-facing `LICENSES.md`, `PYODIDE-LICENSE.txt` and `licenses/<package>/...` beneath `public/runtime/pyodide/<version>/`. These outputs are ignored build artifacts. The deployed notice page is available at `runtime/pyodide/<version>/LICENSES.md` and is linked from the Offline availability panel.

The generated package table records each locked wheel's version, declared license, upstream source URL and included license texts. When a source wheel omits a full license text, the notice identifies that omission and links its upstream project. Pyodide's MPL 2.0 text is separately pinned by URL, byte count and SHA-256 in the runtime lock. Update the lock and review the generated notice output together when changing runtime or package versions.

The Pyodide release documentation identifies the same-origin deployment distribution and versioned asset set; its 314.0.8 documentation says the project uses Mozilla Public License 2.0: [Pyodide license and release](https://pyodide.org/en/stable/project/about.html), [distribution files](https://pyodide.org/en/stable/usage/downloading-and-deploying.html). The release's license file is downloaded from its version tag and hash-checked during preparation.

The production build prints the measured total artifact size, an aggregate gzip size, and manifest counts. Do not commit staged `public/runtime` contents or generated `dist/`. CI and Pages run `npm ci` followed by `npm run check` so the checked-in lockfiles stay the source of truth.
