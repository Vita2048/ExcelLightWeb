# application

Session, commands, atomic history, ports and detached mixed-generation orchestration. Sole authored-state writer. A mixed result is accepted against its captured basis as one session commit, so intermediate native/Python caches never become visible or saved.

`WorkbookSession` accepts a new or detached workbook and an injected ID factory. Use `capture` before detached I/O or structural planning. Use `tryExecute` or `tryAdopt` to reject stale proposals, and pass the original captured snapshot to `markSaved` after storage confirms success. Native cache message types are in `workbook-protocol.ts`.

`editing-commands.ts` plans edits from a captured workbook. A clipboard move combines source removal, destination writes, dependent source rewrites, names and merges in one session transaction.
