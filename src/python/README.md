# python

S02 contains the pinned runtime lock, module worker client and Python transport adapter for the feasibility suite. The probe receives explicit worksheet snapshots and runs the existing demo scripts with an isolated namespace per run. [Development commands](../../docs/Development.md#browser-python-feasibility)

S20 adds `references.ts`, a syntax-aware lexer and rewrite service for literal `xl("...")` calls and workbook names. It respects comments, raw/triple-quoted strings and escapes and deliberately ignores dynamic or aliased `xl` calls. Copy, cut, Fill, sort, sheet/axis structure edits and name/table renames use it to translate source and legacy bindings; the same module builds Monaco xl caret insertion input through `src/features/python`.

S20 owns workbook execution, local consent, runtime preflight and atomic spill application. Mixed recalculation remains S22; figure placement and recovery remain S23. The probe uses JSON and base64 to measure compatibility before choosing the production transport. Its Python worker is for trusted scripts and has the browser capabilities of its origin.
