# Assignment Null-choice module queries

Bare Choose, Switch and IIf calls whose literal arguments return Null check whether any top-level module declaration has the same name. Previously every eligible assignment scanned all module children. The rule now lazily caches the exact queried predicates, including missing names, across procedures within a single invocation. Only these three names can be queried; no complete index or persistent metadata cache is added. Qualified VBA calls bypass the query.

This guard intentionally has narrower scope than the runtime name scope. Its existing behavior considers any direct module declaration, regardless of case or kind; local/project membership is not substituted. The cache refreshes for each invocation, including reused bound roots and changed conditional activity. Messages, ordering and exact diagnostic spans remain unchanged.

## Reproduction and measurements

Run from the repository root with installed dependencies:

```powershell
node scripts/benchmark-assignment-choice-shadow.mjs --baseline=9c5adbb9
node scripts/benchmark-assignment-choice-shadow.mjs
```

The baseline substitutes only assignments.ts from that commit into the current analyzer harness. Fixtures contain the listed number of module variables and 1,000 assignments to a Long. Bare expressions are Choose(0, 1), Switch(False, 1) and IIf(False, 1, Null). The qualified control uses VBA.Choose(0, 1), and the scalar control assigns 1. The benchmark checks 1,000 diagnostics for every call fixture and none for the scalar control.

Parsing and binding are outside timing. The parsed module is reused; each sample gets a fresh bound root, with an identity assertion, so symbol-keyed name/type caches start fresh. This measures checkAssignmentTypes, rather than complete analyzer or editor latency, and does not claim cold parsed-body dataflow measurements. Three warmups precede 15 samples. Before and after ran sequentially without concurrent audit tests, probes or benchmarks.

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor; measured locally on 2026-10-03. All times in milliseconds.

| Declarations / mode | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: |
| 10-choose | 2.690 | 2.526 | 3.348 | 3.398 |
| 10-switch | 2.294 | 2.247 | 2.556 | 2.522 |
| 10-iif | 3.127 | 3.188 | 3.563 | 3.599 |
| 10-qualified | 36.723 | 35.956 | 61.360 | 58.395 |
| 10-scalar | 1.639 | 1.648 | 2.250 | 2.291 |
| 100-choose | 2.538 | 2.062 | 3.219 | 2.663 |
| 100-switch | 2.390 | 1.933 | 3.080 | 2.728 |
| 100-iif | 3.243 | 2.688 | 4.349 | 3.455 |
| 100-qualified | 37.446 | 38.054 | 38.148 | 40.651 |
| 100-scalar | 1.552 | 1.494 | 2.242 | 2.041 |
| 1000-choose | 6.742 | 2.156 | 7.609 | 3.063 |
| 1000-switch | 6.646 | 2.072 | 7.038 | 2.957 |
| 1000-iif | 7.559 | 2.938 | 11.359 | 5.705 |
| 1000-qualified | 71.462 | 72.879 | 81.761 | 75.648 |
| 1000-scalar | 1.651 | 1.570 | 3.176 | 2.901 |
| 3000-choose | 17.737 | 2.419 | 22.996 | 3.761 |
| 3000-switch | 17.151 | 2.370 | 17.673 | 3.866 |
| 3000-iif | 18.227 | 3.174 | 21.074 | 4.519 |
| 3000-qualified | 200.350 | 199.839 | 235.670 | 250.404 |
| 3000-scalar | 1.905 | 1.907 | 3.074 | 3.240 |

With 3,000 declarations, medians improve about 7.3 times for Choose, 7.2 times for Switch and 5.7 times for IIf. Small controls have mixed differences, including some regressions; qualified calls and scalar assignments do not request the cache. The qualified control has substantial unrelated rule/dataflow work in both versions. This change removes repeated module-name scans, not all remaining assignment costs.

Separate getter instrumentation on actual top-level symbol names, with 100 declarations and 100 assignments, recorded:

| Expression | Before name reads | After name reads | Diagnostic count |
| --- | ---: | ---: | ---: |
| Bare Choose | 10,604 | 605 | 100 |
| Bare Switch | 10,604 | 605 | 100 |
| Bare IIf | 10,604 | 605 | 100 |
| Qualified Choose | 504 | 504 | 100 |
| Scalar | 504 | 504 | 0 |

These counts include all top-level name reads in the rule, not only this helper. Counters run outside timings and are not added to production.

## Validation

- Type check passed.
- Focused assignment tests: 172 passed.
- Full suite: 616 files, 12,746 tests passed, 13 skipped.
- Seven regression tests cover name-read bounds across two procedures, module shadow kind/case, qualified spans, existing local-shadow semantics, reused-root mutation and conditional declaration refresh.
- 1,500 differential cases preserve complete analyzer outputs across bare/qualified functions, literal outcomes, malformed input, shadows, conditional compilation, multiple procedures, branches/loops and scalar/Variant target types.

The generated differential harness and instrumentation remain audit scratch material. The timing benchmark and regression tests are committed.
