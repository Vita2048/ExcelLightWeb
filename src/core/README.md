# core

Pure document types, validation, identity, styles, references and formats. No UI/browser imports.

`workbook.ts` defines the plain-DTO model and sparse lookup. `workbook-validation.ts` validates detached workbooks. `workbook-delta.ts` compares authored state and applies or inverts deltas produced from validated snapshots. The application session owns live state.
