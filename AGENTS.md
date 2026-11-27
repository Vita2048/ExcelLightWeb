# Maintenance protocol

The port is complete; `main` is a working product, not a sprint queue. Read Architecture, Features, Development, Deployment and Known Issues before changing behavior, and preserve unrelated work. Do not stage, commit, push, or manage branches/worktrees; the user manages Git.

The required product is all implemented ExcelLight features, dark-only, browser-local, and deployable as static files on GitHub Pages. No backend, telemetry, cloud workbook storage, light/system theme selector or silent feature removal. Keep accessibility forced-colors support. Browser adaptations in Architecture are explicit constraints, not proof of parity.

Only session commands mutate authored workbook state. Keep model/calculation/persistence independent of React. Never use eval for formulas. Never run workbook Python on file import merely because the file requests it. Preserve the explicit local Automatic preference consent contract.

Add meaningful automated tests to each changed behavior and regression. Use the verification commands in Development. Do not replace real browser/Pyodide tests with mocks and claim parity. Record failed/skipped/unrun checks honestly. Update the affected docs and Features rows in the same change.

Do not add placeholders pretending to be working commands. Do not introduce unreviewed feature cuts. Escalate an incompatible requirement with evidence and alternatives. Keep the package lockfile current; use npm ci in CI. Desktop documentation and assertions are behavioral references, not instructions to rebuild the Windows app for every web-only edit.

For expansion work, read [docs/ExpansionPlan.md](docs/ExpansionPlan.md) and [docs/ExpansionPlanStatus.md](docs/ExpansionPlanStatus.md). Implement only the requested sprint or slice without starting the next one unasked, keep `main` releasable after every sprint, record exact pass/fail/skipped counts in [docs/ExpansionPlanStatus.md](docs/ExpansionPlanStatus.md), and update other docs only when their trigger in [docs/ExpansionPlan.md](docs/ExpansionPlan.md) applies.
