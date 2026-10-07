# Event completion prefix gates

Event completion ran on the foreground completion path even when the typed
prefix could not name a handler. It built every dynamic form/control definition,
parsed the module, gathered procedures and checked scope before filtering names.

Filter handler definitions by the current prefix before requesting source facts.
VB6 and Access owners also skip event metadata when the prefix cannot overlap
`owner_`: either the owner prefix starts with the typed word, or the typed word
starts with the owner prefix. Full handler filtering remains necessary afterwards.
Empty prefixes still enumerate all events. Existing-handler exclusion and body
refusal still run for matching prefixes. Context and designer arrays are read
fresh; no cache or public API was added.

## Reproduction

```powershell
npx vitest run tests/eventHandlerPrefixWork.test.ts tests/vbaEventHandlerCompletion.test.ts tests/diagnostics/eventHandlerSignatures.test.ts
node scripts/benchmark-event-prefix-gates.mjs --baseline=099741b137d66b328583fb5df5f768f10794ac07
node scripts/benchmark-event-prefix-gates.mjs
```

Eleven new work regressions fail on the baseline: nine unmatched-prefix cases
across three EOLs and static/VB6/Access contexts still parse, and two matching
control cases query 1,001 event owners instead of one. The fixed unmatched cases
perform no parse or event enumeration; matching control queries enumerate one
owner and return the complete independently expected stub. Another control
preserves existing-handler exclusion and sees a newly added designer control.

## Measurement

Node 24.18.0, Ryzen 7 9800X3D. Nine samples after three warmups, process order
baseline/current/current/baseline. Static cold samples include parser/lexer work
on fresh snapshots; warm samples prime parsing. Dynamic form samples use a
20-query batch average. Source/model creation and complete expected-output
assertions are outside the clock. This measures the actual event resolver,
not the whole completion provider or editor.

| 1,000-procedure/control workload | Baseline median ms/query | Fixed median ms/query |
| --- | ---: | ---: |
| Static unmatched, cold source | 3.8613–3.9791 | 0.0150–0.0157 |
| Static unmatched, warm source | 0.0599–0.0605 | 0.0066–0.0072 |
| Static matching, warm source | 0.0585–0.0589 | 0.0779–0.0830 |
| VB6 unmatched | 0.167065–0.17334 | 0.019905–0.02636 |
| VB6 last control | 0.17538–0.18172 | 0.022145–0.040245 |
| Access unmatched | 0.17856–0.190705 | 0.019805–0.01995 |
| Access last control | 0.18542–0.188955 | 0.021995 |
| VB6 blank prefix, all events | 0.176775–0.17898 | 0.18112–0.182045 |
| Access blank prefix, all events | 0.189625–0.215785 | 0.19162–0.196255 |

Matching static warm queries add about 19–25 microseconds in these runs. Blank
form enumeration is effectively unchanged or slightly slower and still creates
all output rows. Cold matching queries still parse; their measured medians are
not evidence of a parser improvement. Fixed cold negative outliers reached
0.6883 ms, so no tail-latency or whole-editor guarantee is claimed.

## Validation

Types and 52 focused tests pass. Full suite: 742 files, 14,598 tests, 33 skipped,
82.19 seconds. Frozen AST/token/trivia and designer-context baseline comparison
covers 8,256 oracle sources and 480 generated prefix/declaration/body snapshots.
All 183,456 complete completion arrays match (1,081,650 rows). Controls include
blank/partial/full/Unicode owner prefixes, control arrays, MDI forms, Access
forms/reports/section event classes, existing handlers and open procedure bodies.
