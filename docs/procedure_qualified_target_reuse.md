# Qualified procedure target reuse

Introduce Parameter's query-local call binder searched every owning-module declaration again for each accepted owning-module-qualified call. The receiver must be resolved at each call site because locals and parameters can shadow a module qualifier, but the changed declaration itself is fixed for the query.

Resolve that declaration lazily once inside the binder, including a missing-result sentinel. Keep receiver binding and caller/private visibility checks at every call site. Bare calls and rejected qualifiers still avoid the lookup. A new refactor request constructs a new binder, so changes to source or visibility cannot retain the old declaration. No global or persistent cache is added.

Eight new tests cover 1/1,000 public/private qualified calls, missing targets, complete edit outputs, bracketed/case-insensitive module qualifiers, parameter-shadowed receivers, same-module private calls and fresh Public/Private/Public requests across LF/CRLF/CR. Three work checks fail on baseline: 1,000 lookups instead of one for public and private targets, and two instead of one for a missing target. Five independent output/control cases already pass. Types, 164 focused checks and the full suite pass: 14,820 tests, 33 skipped, 763 passing files and seven skipped files, 73.13 seconds.

All 24,768 complete refactor results match baseline across 8,256 unfiltered oracle sources and three newline styles. Each corpus caller includes injected repeated qualified, bracketed and bare calls to the changed target. This establishes compatibility for the exercised inputs, not validity of every malformed corpus source in Office.

## Measurement

Run `node scripts/benchmark-procedure-qualified-target.mjs` in baseline/current/current/baseline order, using `--baseline=878faab4b92f9e45b0f022cfb59dff04ea5a50fa` for baseline runs. Node 24.18.0, Ryzen 7 9800X3D, three warmups and nine measured rounds, five complete refactor queries per round.

Source construction and parser warming are outside the clock. Query-owned project construction, source binding and edit construction remain inside. The benchmark checks result titles, destination module names and independently expected applied outputs after the clock. All baseline/current output controls agree. No cold-parser, retained-heap, editor-command or renderer-paint improvement is claimed.

Ranges below are medians from two independent runs, milliseconds per complete refactor query. The substantial repeatable gain is the 1,000-qualified-call/1,000-preceding-declaration case; single-call, tiny-owner and bare-call controls are mixed, including slower samples.

| Preceding declarations | Calls | Kind | Baseline ms | Repair ms |
| --- | --- | --- | --- | --- |
| 0 | 1 | qualified | 0.03340–0.03494 | 0.03422–0.03502 |
| 0 | 1 | bare | 0.02584–0.02796 | 0.02468–0.03152 |
| 0 | 1 | private | 0.01970–0.02204 | 0.01982–0.02192 |
| 0 | 1000 | qualified | 1.07592–1.25846 | 1.06136–1.07590 |
| 0 | 1000 | bare | 0.79766–0.87142 | 0.77610–0.83880 |
| 0 | 1000 | private | 0.89474–0.89578 | 0.87228–0.88010 |
| 1000 | 1 | qualified | 0.31858–0.33032 | 0.31658–0.33610 |
| 1000 | 1 | bare | 0.26994–0.41558 | 0.33250–0.45486 |
| 1000 | 1 | private | 0.31646–0.32208 | 0.28516–0.34722 |
| 1000 | 1000 | qualified | 2.20630–2.24854 | 1.14484–1.16728 |
| 1000 | 1000 | bare | 0.97936–1.02594 | 0.97222–0.98392 |
| 1000 | 1000 | private | 2.18504–2.20338 | 1.07946–1.11800 |
