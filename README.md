# ExcelLight Web

Browser-only port of `C:\Temp\ExcelLight`, with a dark UI and GitHub Pages hosting. Workbooks and Python computation stay on the user's device. This repository currently contains **planning documents and a runnable scaffold**, not the implemented spreadsheet.

## Start here

- [Implementation plan and sprint sequence](docs/Implementation-Plan.md)
- [Architecture and technology decisions](docs/Architecture.md)
- [Feature parity matrix](docs/Feature-Parity.md)
- [Development and test strategy](docs/Development.md)
- [GitHub deployment](docs/Deployment.md)
- [Sprint handover template](docs/templates/Sprint-Handover.md)
- [Agent instructions](AGENTS.md)

## Local commands

Use Node 22.12+ in the Node 22 release line and npm. The lockfile pins exact dependency resolutions.

```powershell
npm ci
npm run dev
npm run check
npx playwright install chromium firefox webkit
npm run test:e2e
```

`test:e2e` requires the production build from `check`. It starts and stops its own local static server. `npm run preview` serves the production build for manual inspection. No backend, database service, API keys, or cloud account is needed for local development.

## Agent starting prompt

> Read AGENTS.md, docs/Architecture.md, docs/Feature-Parity.md and docs/Implementation-Plan.md. Implement S01 using docs/sprints/S01.md. Complete its acceptance checks, update the parity matrix and write docs/handovers/S01.md using the handover template. Report checks actually run and blockers. Do not claim later sprints are implemented.

Replace S01 with the next sprint only after its prerequisites are completed. The scaffold does not mean S01 is complete; S01 freezes the source baseline and fixtures.
