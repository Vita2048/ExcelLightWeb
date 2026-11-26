# Implementation plan

## Goal and baseline

Port all implemented features of C:/Temp/ExcelLight into this directory, with a dark-only UI, browser-local workbooks and computation, GitHub source control and static GitHub Pages deployment. Scope means the desktop application's documented subset, not every Excel feature. CSV, macros, pivots, collaboration and general dynamic arrays were not implemented in the source and are not new requirements.

This deliverable contains the planning set, directory boundaries, React shell, strict TypeScript configuration, automated smoke tests and CI/deployment workflows. No spreadsheet, Python or persistence sprint is complete yet. Do not confuse source desktop sprint numbers with these new web sprint IDs.

## How to execute

Use one slice as one sprint. S01 freezes the observable baseline; S02/S03 retire the biggest feasibility risks before committing to a full port. Then follow numerical order S04–S26. Dependencies below also document the minimum prerequisites, but numerical order is the recommended single-agent workflow. Each sprint is a usable or verifiable increment, ends with tests and a handover, and must not leave simulated working features in the UI.

## Sprint sequence

| Sprint | Deliverable | Prerequisites |
| --- | --- | --- |
| [S01](sprints/S01.md) | Baseline and test fixtures (complete) | none |
| [S02](sprints/S02.md) | Browser Python feasibility gate (complete) | S01 |
| [S03](sprints/S03.md) | Formula and file adapter feasibility gate (complete) | S01 |
| [S04](sprints/S04.md) | Workbook model, session and history | S01, S03 |
| [S05](sprints/S05.md) | Dark shell and virtual worksheet | S04 |
| [S06](sprints/S06.md) | Editing, clipboard and fill | S05 |
| [S07](sprints/S07.md) | Structural edits and dimensions | S06 |
| [S08](sprints/S08.md) | Formula language and functions | S03, S07 |
| [S09](sprints/S09.md) | Native dependency engine and formula UI | S08 |
| [S10](sprints/S10.md) | Number, date and special formats | S09 |
| [S11](sprints/S11.md) | Fonts, fills and Format Painter | S10 |
| [S12](sprints/S12.md) | Borders and drawing tools | S11 |
| [S13](sprints/S13.md) | Alignment and merged ranges | S12 |
| [S14](sprints/S14.md) | Names and tables | S13 |
| [S15](sprints/S15.md) | Sort, filter and statistics | S14 |
| [S16](sprints/S16.md) | Find, view and navigation polish | S15 |
| [S17](sprints/S17.md) | Native .exlt persistence | S03, S16 |
| [S18](sprints/S18.md) | Browser file lifecycle and local recovery | S17 |
| [S19](sprints/S19.md) | XLSX exchange and loss reports | S03, S18 |
| [S20](sprints/S20.md) | Python authoring, references and manual execution/results | S02, S19 |
| [S21](sprints/S21.md) | Folded into S20 by user-approved scope expansion | S20 |
| [S22](sprints/S22.md) | Mixed native/Python recalculation | S09, S20 |
| [S23](sprints/S23.md) | Figures, placements and redraw | S22 |
| [S24](sprints/S24.md) | Offline packaging and hardening | S23 |
| [S25](sprints/S25.md) | Cross-browser parity and performance | S24 |
| [S26](sprints/S26.md) | Release and GitHub Pages acceptance | S25 |

## Milestones and gates

- S01–S03: baseline and evidence that browser Python, formulas and file adapters are viable. Failed feasibility blocks dependent implementation and triggers a documented architecture decision.
- S04–S16: spreadsheet model and complete native editing/calculation/formatting/data/view workflows.
- S17–S19: desktop native-file interoperability, browser save/recovery and reviewed XLSX exchange.
- S20–S23: source authoring and manual Python results, mixed recalculation and linked figures.
- S24–S26: offline packaging, all-feature acceptance and actual GitHub Pages release.

## Shared Definition of Done

1. All sprint acceptance cases implemented and backed by relevant tests; original source behavior and approved browser adaptations are traceable in Feature-Parity.
2. Typecheck, unit/integration tests, production build and applicable browser/runtime/parity suites pass. Report exact commands and failing/skipped/unrun checks. No skipped test is evidence of parity.
3. Invalid input and cancellation preserve the document; no stale async result mutates the wrong session. Authored state, history, save identity and consent remain correct.
4. Dark-only UI, keyboard/focus behavior, local data constraint and repository-subpath hosting still hold. Review new ribbon commands, flyouts and icons against the desktop in their functional sprint. S25 reviews the complete ribbon across browsers.
5. Update affected architecture/development/deployment/compatibility docs; record dependency choices and actual runtime versions. No introduced critical unresolved regression.
6. Save a concise handover in docs/handovers/SNN.md, including next action and known risks. Mark the sprint complete only when its required evidence exists.

## Folder structure

```text
ExcelLightWeb/
  .github/workflows/      # CI and Pages deployment on main pushes
  config/                # build/deployment configuration helpers
  docs/
    Architecture.md      # decisions, contracts, browser adaptations
    Feature-Parity.md    # source-to-sprint/test traceability
    Implementation-Plan.md
    Development.md
    Deployment.md
    source-baseline/     # S01 snapshot hashes and fixture provenance
    sprints/S01..S26.md  # one implementation specification per sprint
    handovers/           # sprint results, including initial scaffold report
    templates/           # simple handover template
  public/runtime/        # ignored, versioned Pyodide assets staged by the S20 production build
  scripts/               # build/test/runtime acquisition helpers
  src/
    app/                 # composition root and shell
    core/                # UI-free document/types/validation/references
    application/         # commands/history/session/ports/orchestration
    calculation/         # formula AST/functions/dependencies
    persistence/         # native, XLSX, browser storage and file adapters
    python/              # Pyodide integration, xl discovery, DTO conversion
    workers/             # calculation/Python/I/O worker entrypoints
    grid/                # canvas geometry/rendering/input/accessibility
    features/            # ribbon/dialog/editor controllers by feature
    styles/              # dark tokens and shared styles
  tests/
    unit/                # pure semantic behavior
    integration/         # transactions/adapters/worker protocol
    parity/              # desktop differential oracle suites
    fixtures/            # small approved samples and manifest
    e2e/                 # actual browser user journeys/runtime tests
    performance/         # reproducible workload measurements
  index.html
  package.json / package-lock.json
  tsconfig.json / vite.config.ts / vitest.config.ts / playwright.config.ts
```

Directories carry boundary READMEs until their sprint implements code. This is a single application repository, not an unnecessary multi-package monorepo. Keep public runtime artifacts reproducible and out of source history; the acquisition script and checksum manifest belong in Git. User workbooks must not be copied into public assets.

