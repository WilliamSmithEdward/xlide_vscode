# Event completion boundaries during parser recovery

Missing procedure terminators can leave many open procedure nodes. Event-handler
completion checked each open node's extent by scanning every module member for
the next greater start offset. A recovered module with 1,000 open procedures
performed about 1.5 million start reads in one query.

The first open boundary still uses a single member scan, avoiding an index for
one open procedure or an early body hit. If another boundary is needed, the
request builds one sorted list of member starts and uses strict-successor binary
search thereafter. The sort uses a new array, preserving shared parser nodes.
Duplicate and unordered starts retain the old minimum-greater-start semantics,
and boundaries remain capped at module end. Closed procedures keep their own
inclusive spans. No source/project cache or public API was added.

## Reproduction

```powershell
npx vitest run tests/eventHandlerOpenProcedureWork.test.ts tests/vbaEventHandlerCompletion.test.ts
node scripts/benchmark-event-open-procedures.mjs --baseline=1b445e849ef18382680c4d7bb71a49de7814447e
node scripts/benchmark-event-open-procedures.mjs
```

Nine work-count regressions use 10/100/1,000 missing `End Sub` headers across
LF, CRLF and CR. They verify the real parser's open-node count and the complete
expected following `Workbook_Open` stub. All fail on the baseline and pass with
a bound of six start reads per member (at most 6,012 for the largest case).
Three controls preserve refusal inside first/last open bodies and allow the
following closed boundary.

## Measurement

Node 24.18.0, Ryzen 7 9800X3D, nine samples after three warmups; each sample runs
20 actual completion queries and reports milliseconds per query. Separate
process order baseline/current/current/baseline. Parsing is primed before the
clock; boundary checks, existing-name/definition filtering and stub construction
are timed. Complete expected outputs are independently asserted.

| 1,000-procedure shape | Baseline median ms/query | Fixed median ms/query |
| --- | ---: | ---: |
| Many open procedures, caret after a closed procedure | 0.89043–0.90277 | 0.101575–0.10771 |
| One open procedure | 0.05267–0.05271 | 0.051705–0.052345 |
| All closed | 0.067745–0.07111 | 0.06599–0.06953 |
| Caret in the first open body | 0.008675–0.008715 | 0.00865–0.011575 |

This removes quadratic boundary work in recovered source. Early body and small
cases remain cheap and timings are mixed; it does not establish faster parsing
or a whole-editor latency improvement. The index's construction/sort is timed,
and requests still gather/filter procedure names after module-level detection.

## Validation

Types and 26 focused tests passed. A frozen AST/token/trivia baseline comparison
covers 8,256 oracle sources plus 27 generated malformed modules, three document
contexts and three caret positions: all 74,547 complete completion arrays match
(390,003 rows). A separate 10,000-case frozen boundary differential includes
unordered and duplicate starts, mixed open/closed nodes, varied caret offsets
and module ends; every decision matches. Full suite passed: 740 files, 14,587 tests, 33 skipped tests, in 84.87 seconds.
