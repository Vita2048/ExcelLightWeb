# Implementation protocol

Build the port one approved sprint at a time using `docs/sprints/SNN.md`. Read the latest handover, Architecture, Feature-Parity, Development and Deployment first. Preserve unrelated work. Keep the desktop source read-only. Do not stage, commit, push, or manage branches/worktrees; the user manages Git.

The required product is all implemented ExcelLight features, dark-only, browser-local, and deployable as static files on GitHub Pages. No backend, telemetry, cloud workbook storage, light/system theme selector or silent feature removal. Keep accessibility forced-colors support. Browser adaptations in Architecture are explicit constraints, not proof of parity.

Only session commands mutate authored workbook state. Keep model/calculation/persistence independent of React. Never use eval for formulas. Never run workbook Python on file import merely because the file requests it. Preserve the explicit local Automatic preference consent contract.

Add meaningful automated tests to each implemented behavior and regression. Use the verification commands in Development. Do not replace real browser/Pyodide tests with mocks and claim parity. Record failed/skipped/unrun checks honestly. Update the affected docs and parity rows. Each sprint ends with `docs/handovers/SNN.md`; a blocked sprint stays blocked and dependent work does not proceed on an unproven assumption.

Do not add placeholders pretending to be working commands. Do not introduce unreviewed feature cuts to pass a sprint. Escalate an incompatible requirement with evidence and alternatives. No subagents or owner fields are required. Keep the package lockfile current; use npm ci in CI. Source application docs are behavioral references, not instructions to build the unchanged Windows app for every web-only edit.
