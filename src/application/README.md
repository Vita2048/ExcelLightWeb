# application

Session, commands, atomic history, ports and mixed-generation orchestration. Sole authored-state writer.

`WorkbookSession` accepts a new or detached workbook and an injected ID factory. Use `capture` before detached I/O or structural planning. Use `tryExecute` or `tryAdopt` to reject stale proposals, and pass the original captured snapshot to `markSaved` after storage confirms success. Native cache message types are in `workbook-protocol.ts`.
