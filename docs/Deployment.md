# GitHub repository and GitHub Pages deployment

## Compatibility

The planned application is static: GitHub stores source, Actions installs/tests/builds, and Pages serves `dist/`. Workbooks and computation stay in the browser. There is no server API, server-side rendering, secret, cloud database or uploaded workbook. Pyodide/Monaco/runtime assets must be bundled with the build in their implementation sprints.

The app uses relative asset URLs (`base: './'`), so both `https://OWNER.github.io/REPOSITORY/` and a domain-root site work. Keep navigation on the entry page or use hash routes. Do not add history routing that requires a server fallback. VITE_BASE_PATH is an optional process environment override for a slash-delimited explicit base; default relative mode is what the checked-in browser tests exercise. Custom overrides need matching deployment tests.

Workers use `new Worker(new URL(..., import.meta.url), { type: 'module' })`; runtime assets resolve against the configured build base; service worker registration uses the app base and scope. Never hardcode `/assets`, `/runtime` or a localhost API. No SharedArrayBuffer or custom COOP/COEP headers are required by this design. Stop uses worker termination.

The S09 production browser suite loads the native calculation worker from the built assets at the repository subpath. Vite packages it with the worksheet. Native calculation needs no separate runtime asset staging. S10 bundles the format engine and locale catalog into the application and worker. Formatting makes no network requests. S11 uses browser-local fonts and canvas for font and fill previews. The app needs no hosted font files. S13 serves the original desktop AppIcon.ico as a static favicon through the configured asset base. Browser tests request it at both supported deployment paths.

## First publication (user-managed Git)

1. Create a GitHub repository with a main branch and add this directory's contents, including package-lock.json, docs and .github workflows. Do not add node_modules, dist, private workbooks or generated runtime downloads. This task has not created a repository, committed or pushed anything.
2. Select repository **Settings → Pages → Source → GitHub Actions**. Ensure Pages is available for the selected repository/account visibility. Review who can read the published application separately from who can read the repository.
3. Run **Verify** and resolve failures. The workflow runs on pull requests and pushes to main.
4. In **Actions → Deploy GitHub Pages**, select **Run workflow** on main. The workflow installs with npm ci, typechecks/tests/builds, runs Chromium/Firefox/WebKit tests, uploads dist and deploys through the github-pages environment.
5. Record the returned hosted URL and source revision in the release handover. Verify a reload at the repository URL, worker/runtime requests, local save/export/reopen and offline-ready behavior. The current build supports worksheet and hosting checks. Product release acceptance belongs to S26.

Deployment is manually triggered to avoid publishing partial sprint work on every push. It grants contents:read globally and pages:write/id-token:write only to the deployment job; main is required for deployment. The github-pages environment can hold GitHub-native review rules if desired. No personal deployment token is needed. Failed checks prevent deployment.

## Runtime packaging before product release

S02 records the pinned Pyodide release and package closure in [the probe lock](../src/python/runtime-lock.json). Its acquisition script verifies SHA-256 and serves the assets in a dedicated test build. S24 will add license notices, production asset staging and CI download steps before the Pages build.

Check total site/runtime sizes and individual assets against current [Pages limits](https://docs.github.com/en/pages/getting-started-with-github-pages/github-pages-limits). Browser asset requests naturally go to the host; application workbook data does not. User Python can make requests, so local computation is not a promise that arbitrary user scripts cannot use the network.

## Custom domain, upgrades and rollback

An optional custom domain uses the same static build; configure DNS and Pages and add the appropriate public/CNAME only when a domain is chosen. It is not required now. Browser data is origin scoped: switching to a custom domain does not migrate workbooks. Export/import before switching. Project paths on one github.io origin are not separate security boundaries.

S24 must defer service-worker activation/reload while edits are dirty, keep schema migrations transactional and test old-cache/new-code interactions. Store local recovery before asking users to refresh. Never promise first-visit offline support.

Rollback by redeploying a previously verified revision through the same workflow, with user-managed Git selecting that revision on main if necessary. Record the artifact/runtime manifest that worked. An older app may not read a newer local DB schema: retain reversible/backed-up migrations, or export with the newer version before rollback. Do not delete IndexedDB as a rollback shortcut.

## References

[Vite's official static deployment guide](https://vite.dev/guide/static-deploy) documents Pages and base paths. [Pyodide workers](https://pyodide.org/en/stable/usage/webworker.html) and [interrupt requirements](https://pyodide.org/en/stable/usage/keyboard-interrupts.html) explain the worker design. References checked 2026-11-17; recheck action/runtime compatibility before S26 release.
