# Initial planning and scaffold handover

Status: complete for the requested planning/scaffolding task; S01–S26 remain planned.
Date: 2026-11-17

## Delivered

- Architecture and stack recommendation based on the desktop docs and selected source contracts. Browser adaptations and Python/file feasibility gates are explicit.
- Feature parity matrix, 26 dependency-ordered sprint specifications with acceptance tests, directory boundaries, agent instructions and handover template.
- React/TypeScript/Vite dark scaffold, dependency lockfile, Vitest and Playwright setup, CI and manually triggered GitHub Pages workflow.
- SHA-256 planning inspection manifest for 12 desktop authority files. This is not the complete S01 fixture/oracle capture.

## Verification

| Command / scenario | Result | Evidence |
| --- | --- | --- |
| npm install with trusted system CA | Pass | 49 packages added; audit reported zero vulnerabilities at installation |
| npm run check | Pass | TypeScript, 6 configuration tests, production Vite build |
| npm run test:e2e | Pass | 6 tests: root and repository subpath on Chromium, Firefox and WebKit; OS light still renders dark |
| Local Markdown link check | Pass | All relative links resolved; 26 sprint specifications found |
| Desktop app / full port parity / real Pyodide | Not run | No port functionality implemented; assigned to future sprints |
| GitHub Actions / hosted deployment | Not run | No repository publication performed |

Browser tests initially could not reach loopback sockets in the restricted sandbox; running with the approved external execution mode passed. Dependency installation initially encountered a certificate-chain error; using Node's trusted system certificates resolved it without disabling TLS verification.

## Remaining risks

Python package/API equivalence, browser memory, extended XLSX styles, formula semantics and figure recovery require S02/S03 evidence. Arbitrary desktop Python OS integration cannot be promised in a browser. Local data persistence and download behavior differ from atomic Windows saves; Architecture documents the contract.

Git metadata in the source directory was inaccessible during inspection; the planning manifest records working-tree file hashes instead. S01 must complete the baseline inventory and oracle capture. Source files were not modified. No Git commit/push or deployment was performed.

## Next sprint

Next: S01 — Baseline and test fixtures.
First action: read ../sprints/S01.md and expand ../source-baseline/planning-inspection.json into the command inventory and approved fixture manifest. Use ../templates/Sprint-Handover.md at completion.
