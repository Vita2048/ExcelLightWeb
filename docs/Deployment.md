# GitHub repository and GitHub Pages deployment

## Compatibility

The application is static: GitHub stores source, Actions installs/tests/builds, and Pages serves `dist/`. Workbooks and computation stay in the browser. There is no server API, server-side rendering, secret, cloud database or uploaded workbook. Pyodide, Monaco and runtime assets are bundled with the build.

The app uses relative asset URLs (`base: './'`), so both `https://OWNER.github.io/REPOSITORY/` and a domain-root site work. Keep navigation on the entry page or use hash routes. Do not add history routing that requires a server fallback. `VITE_BASE_PATH` is an optional process environment override for a slash-delimited explicit base; default relative mode is what the checked-in browser tests exercise. Custom overrides need matching deployment tests.

Workers use `new Worker(new URL(..., import.meta.url), { type: 'module' })`; runtime assets resolve against the configured build base; service-worker registration uses the app base and scope. Never hardcode `/assets`, `/runtime` or a localhost API. No SharedArrayBuffer or custom COOP/COEP headers are required. Stop uses worker termination.

The native calculation worker is packaged with the worksheet and needs no separate runtime asset staging. The format engine and locale catalog are bundled into the application and worker; formatting makes no network requests. Font and fill previews use browser-local fonts and canvas. The desktop AppIcon.ico is served as a static favicon through the configured asset base, and browser tests request it at both supported deployment paths.

## First publication (existing repository; user-managed Git)

1. Use the existing `Vita2048/ExcelLightWeb` repository on `main`; do not create a second repository or copy this directory elsewhere. This checkout's `origin` already points to that repository. Keep `package-lock.json`, docs and `.github` workflows in the repository. Do not add `node_modules`, `dist`, private workbooks or generated runtime downloads. The user manages all Git operations.
2. In repository **Settings → Pages → Source**, select **GitHub Actions**. Ensure Pages is available for the repository/account visibility. Review who can read the published application separately from who can read the repository. The workflow source matters: selecting **Deploy from a branch** publishes repository files as-is and does not run Vite's production build.
3. Run **Verify** on the intended `main` revision and resolve failures. The verification workflow runs on pull requests and pushes to `main`.
4. The Support SLA example in [docs/showcase](showcase) is the only workbook inside the build. It was authored for this repository by the project owner, and its [manifest entry](../tests/fixtures/manifest.json) records that provenance, so publishing it carries no third-party redistribution review. The sales workbook in the same folder is test-only and is not bundled.
5. Push to `main`, or in **Actions → Deploy GitHub Pages** select **Run workflow** on `main`. This project workflow (`.github/workflows/pages.yml`) runs on every push to `main` and on manual dispatch: it installs with `npm ci`, typechecks/tests/builds, runs Chromium/Firefox/WebKit tests, uploads `dist` and deploys through the `github-pages` environment. A failed check prevents deployment. Do not use the default **pages build and deployment** branch workflow for this Vite app: a green run there can still publish unbuilt source files, which leads to `/src/main.tsx` 404s instead of loading the application bundle.
6. The expected project-site URL is `https://vita2048.github.io/ExcelLightWeb/`; `https://github.com/Vita2048/ExcelLightWeb` remains the source repository URL. After a deployment, verify a reload at the hosted URL, worker/runtime requests, local save/export/reopen and offline-ready behavior.

Deployment runs automatically for every push to `main`, and can also be started manually. Because the published site tracks `main`, every push must be releasable: keep partial work on pull request branches (the **Verify** workflow runs there) and merge to `main` only when it is publishable. Each deployment still installs, typechecks, tests and builds from that revision, so a failed check prevents publication. Runs are serialized by the `pages` concurrency group, so rapid pushes deploy in order; manual dispatch retries the current `main` revision without a new push. The workflow grants contents:read globally and pages:write/id-token:write only to the deployment job; `main` is required for deployment. The `github-pages` environment can hold GitHub-native review rules if desired. No personal deployment token is needed.

## Runtime packaging and offline use

The pinned Pyodide release and package closure are recorded in [the runtime lock](../src/python/runtime-lock.json). `npm run build` verifies and stages the locked assets and Pyodide license, generates wheel license files and `LICENSES.md`, then emits a checksum manifest and service worker into `dist/`. Runtime downloads and generated license copies are ignored source artifacts; CI and Pages recreate them from the lock before deployment. The worker install caches the static app assets under the exact `/` or repository-path scope, including the bundled Support SLA example. It does not cache user workbooks or automatically download Python.

On every page load, a worker initializes Pyodide and loads NumPy, pandas and Matplotlib from the same-origin static assets without running workbook Python source. First visit requires connectivity for this initialization. The app's **Prepare Python for offline use** action separately downloads and verifies the full locked Python asset set, including its notices, into Cache Storage. The app reports Python as offline-ready only after every expected file and the final readiness record have been stored. A quota, stale checksum, interrupted download or missing cache leaves it not ready; a retry resumes verified files and reports failures. When the browser is offline and a required runtime file is missing, Python reports an offline cache miss. The production browser tests exercise both root and repository-subpath service-worker scopes, Monaco, offline reopen, and a real Pyodide run from the cache.

The update prompt waits for a user action. Dirty work is placed in the existing browser-local recovery store before the initiating tab reloads; another dirty same-scope tab holds the update until that tab is saved or closed. Existing tabs are not forcibly reloaded. Recovery uses the same IndexedDB lease and transactional native package records as the workbook lifecycle.

File Home bundles the Support SLA example as a static `.exlt` asset. Choosing it loads the fixture into a new workbook identity and leaves it unsaved until the user chooses Save or Save As; no workbook is uploaded.

The build prints measured full artifact size (compressed and uncompressed) and app/runtime manifest counts. Check the actual values and individual assets against current [Pages limits](https://docs.github.com/en/pages/getting-started-with-github-pages/github-pages-limits). Browser asset requests naturally go to the host; application workbook data does not. User Python can make HTTPS/WebSocket requests, so local computation does not prevent user-authored code from networking.

## Custom domain, upgrades and rollback

An optional custom domain uses the same static build; configure DNS and Pages and add the appropriate `public/CNAME` only when a domain is chosen. Browser data is origin scoped: switching to a custom domain does not migrate workbooks, so export/import before switching. Project paths on one `github.io` origin are not separate security boundaries.

Service-worker activation and reload are deferred while edits are dirty; schema migrations stay transactional and are tested for old-cache/new-code interactions. Local recovery is stored before users are asked to refresh. First-visit offline support is never promised.

Rollback by redeploying a previously verified revision through the same workflow, with user-managed Git selecting that revision on `main` if necessary. Record the artifact/runtime manifest that worked. An older app may not read a newer local DB schema: retain reversible/backed-up migrations, or export with the newer version before rollback. Do not delete IndexedDB as a rollback shortcut.

## References

[Vite's official static deployment guide](https://vite.dev/guide/static-deploy) documents Pages and base paths. [Pyodide workers](https://pyodide.org/en/stable/usage/webworker.html) and [interrupt requirements](https://pyodide.org/en/stable/usage/keyboard-interrupts.html) explain the worker design. [GitHub Pages limits](https://docs.github.com/en/pages/getting-started-with-github-pages/github-pages-limits) bound the artifact. Recheck action/runtime compatibility before a major upgrade.
