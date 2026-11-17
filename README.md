# ExcelLight Web

Browser-only port of `C:\Temp\ExcelLight` with a dark UI and GitHub Pages hosting. The current build opens a blank workbook with a virtual worksheet, cell selection, a read-only formula bar, and view controls. Cell editing, file handling, and in-workbook Python execution are scheduled for later sprints.

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
npm run test:worksheet
```

`test:e2e` requires the production build from `check`. It starts and stops its own local static server. `npm run preview` serves the production build for manual inspection. No backend, database service, API keys, or cloud account is needed for local development.

## Agent starting prompt

> For the next approved sprint, read AGENTS.md, the architecture, parity, development and deployment docs, and the latest handover. Implement the approved sprint, run its acceptance checks, update the parity matrix, and write its handover.

The latest handover covers the dark shell and virtual worksheet. [S05 handover](docs/handovers/S05.md)
