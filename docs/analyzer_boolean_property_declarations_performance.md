# Boolean declaration membership for numeric host properties

Numeric host-property checks need to distinguish a Boolean holding True from an ordinary number holding -1. Previously each held numeric value scanned every procedure child to ask whether any matching name had an exact case-insensitive Boolean declared type. The helper now lazily builds the set of direct-child Boolean names once per procedure pass, after a numeric held value first requests it. This bounds distinct-name queries as well as repeated names. Held values remain specific to each statement.

The set preserves ANY matching direct child, including later duplicates and children outside the binding resolver's local-kind filter. It uses exact asType lowercasing without trimming or broader type normalization. The existing binding index cannot substitute directly: it filters children and adds a synthetic return variable. The new set is invocation-local and refreshed when a reused bound root's metadata changes. Literal or nonnumeric values that bypass the known-numeric callback do not initialize it.

## Deterministic cost

Actual symbol name getters measured all reads in checkAssignmentTypes, including other setup/completion work, with 100 unrelated Long locals and 100 Font.Size assignments:

| Value mode | Before name reads | After name reads | Diagnostics |
| --- | ---: | ---: | ---: |
| Held Long 500 | 11,720 | 1,620 | 100 |
| Boolean True | 11,822 | 1,723 | 0 |
| Boolean False | 11,822 | 1,723 | 100 |
| Integer True | 11,822 | 1,722 | 100 |
| Literal 12 | 1,722 | 1,722 | 0 |
| Distinct default Long names | 11,822 | 1,722 | 100 |
| Boolean declared first | 1,822 | 1,723 | 0 |

Counters run separately from timing and are not added to production. Other symbol reads remain; the set indexes declaration membership rather than caching numeric values.

## Reproduction and timings

From the repository root with installed dependencies:

```powershell
node scripts/benchmark-boolean-property-locals.mjs --baseline=9c3bfcf1
node scripts/benchmark-boolean-property-locals.mjs
```

The baseline replaces only assignments.ts from that commit in the current harness. Each fixture has the listed number of unrelated Long locals and 1,000 Range("A1").Font.Size assignments. Held v values are Long 500, Boolean True/False or Integer True. Literal controls assign 12. Distinct fixtures cycle through the unrelated Long locals, whose default zero is refused. The first-Boolean control declares v before the unrelated locals. The benchmark checks 1,000 diagnostics for Long/False/Integer/distinct fixtures and zero for True/literal/first-Boolean fixtures.

Parsing, tokenization and binding are outside timing. Parsed bodies/tokens are reused, with a fresh bound root and fresh completion caches per sample; root identity is asserted. These measurements cover checkAssignmentTypes, not full analyzer/editor latency or a claim about cold parsed-body dataflow. Three warmups precede 15 samples. Before and after ran sequentially after all audit tests and differential probes finished.

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor; measured locally on 2026-10-03. Times in milliseconds.

| Locals / mode | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: |
| 10-long | 5.258 | 5.249 | 6.002 | 6.044 |
| 10-true | 4.085 | 3.934 | 5.191 | 4.666 |
| 10-false | 3.978 | 3.870 | 5.579 | 4.592 |
| 10-integer | 4.834 | 4.700 | 5.589 | 5.530 |
| 10-literal | 3.830 | 3.929 | 4.720 | 5.810 |
| 10-distinct | 4.013 | 3.940 | 4.976 | 5.201 |
| 10-true-first | 3.888 | 3.779 | 4.819 | 4.596 |
| 100-long | 5.551 | 4.696 | 7.918 | 6.962 |
| 100-true | 5.068 | 4.627 | 6.150 | 5.412 |
| 100-false | 5.899 | 4.592 | 7.586 | 5.326 |
| 100-integer | 5.551 | 4.585 | 8.132 | 5.714 |
| 100-literal | 4.461 | 4.549 | 5.223 | 5.448 |
| 100-distinct | 5.267 | 4.708 | 5.774 | 5.441 |
| 100-true-first | 4.667 | 4.467 | 6.422 | 5.347 |
| 1000-long | 17.260 | 11.904 | 18.357 | 13.358 |
| 1000-true | 16.910 | 11.786 | 18.406 | 12.461 |
| 1000-false | 16.927 | 13.046 | 17.413 | 21.300 |
| 1000-integer | 16.606 | 12.576 | 17.123 | 14.867 |
| 1000-literal | 12.016 | 12.457 | 13.480 | 13.107 |
| 1000-distinct | 21.300 | 14.659 | 24.136 | 16.039 |
| 1000-true-first | 13.065 | 12.586 | 15.991 | 19.483 |
| 3000-long | 55.898 | 38.515 | 67.983 | 40.250 |
| 3000-true | 60.298 | 40.024 | 66.425 | 43.493 |
| 3000-false | 55.963 | 39.081 | 59.044 | 43.600 |
| 3000-integer | 56.878 | 38.930 | 65.105 | 40.616 |
| 3000-literal | 40.321 | 38.993 | 49.060 | 51.193 |
| 3000-distinct | 61.239 | 41.757 | 68.919 | 44.601 |
| 3000-true-first | 39.774 | 40.364 | 42.871 | 48.450 |

Large repeated and distinct named-value medians improve roughly 30–34 percent. Literal controls retain substantial unrelated work. Small controls and some p95 results are mixed: the 3,000-local first-Boolean median changes from 39.774 to 40.364 ms, with p95 42.871 versus 48.450 ms. Building a complete set does more first-query declaration inspection than an early successful linear search; subsequent queries are constant-time. This does not establish a universal latency improvement. The set retains only direct Boolean names and is discarded with the procedure pass.

## Validation

- Type check passed.
- Focused property tests: 81 passed.
- Full suite: 619 files, 12,784 tests passed, 13 skipped.
- Twelve regressions cover real read bounds for repeated/distinct/control modes, statement-specific values, ANY duplicate/kind behavior, exact type spelling, mutable bound-root refresh, procedure isolation and conditional declaration activity.
- 1,500 differential cases preserve complete analyzer and direct rule outputs across five Excel property limit families, Boolean/numeric/string/Variant values, duplicate/Static/Const/conditional declarations, updates, separate procedures, With/branches/loops and literal controls.

Counting and generated differential probes remain audit scratch material; the timing benchmark and regression tests are committed.
