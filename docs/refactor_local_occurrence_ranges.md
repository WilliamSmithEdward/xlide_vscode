# Procedure-local identifier occurrence ranges

Inline Variable and Introduce Parameter ask `localUsesIn` for one procedure's
uses and writes. The helper previously found every matching name across the
whole module, then discarded occurrences outside the procedure. Even with
stripped lines cached, the identifier regex rescanned every unrelated line and
allocated matching occurrence records that were immediately discarded.

The source scanner now accepts optional inclusive bounds on occurrence starting
offsets. It binary-searches cached line starts and matches identifiers only on
those physical lines, retaining original line, column, offset and spelling.
`localUsesIn` passes the procedure span directly. The shared whole-source
stripping cache retains continued-comment context, so a slice starting after a
comment continuation cannot turn comment text into a use. Boundary lines are
scanned whole and results filtered by starting offset: words are not truncated
and the previous inclusive-end behavior is preserved. Unbounded callers keep
their existing behavior.

## Validation

Baseline `6059d89dc95f54d986a16eadea77d360859b2a20`: all six new work
regressions fail while independently expected helper results match. With 1,000
unrelated procedures, their body lines receive 2,000 regex calls on the baseline
and zero with the range. LF, CRLF and CR are covered.

Ten new tests cover those work regressions plus Unicode/case matching, multi-name
queries, absolute positions, preceding and internal continued comments,
inclusive bounds inside words, reversed ranges, EOF and unbounded calls.

Type checking and 137 focused tests passed. The full suite passed: 746 files,
14,636 tests; 7 files / 33 tests skipped (139.68 seconds). A baseline differential
across 8,256 oracle sources and 750 generated sources matched 180,120 ranged
queries against independently filtered baseline full-source results, 45,628
complete helper results, and 1,500 complete Inline Variable / Introduce Parameter
results, including refusals.

## Measurement

Run `node scripts/benchmark-refactor-local-occurrences.mjs`, optionally with
`--baseline=6059d89dc95f54d986a16eadea77d360859b2a20`.
Separate baseline/fixed/fixed/baseline processes on Node 24.18.0 / Ryzen 7 9800X3D
used three warmups, nine rounds and batches of ten, reporting per-query averages.
The target procedure is at the beginning, middle or end of a module; unrelated
procedures use the same local name. Parsing and fixture construction are outside
the timer; stripping/token caches are primed by warmups. Complete helper output
is asserted independently outside the timer.

| Unrelated procedures | Position | Baseline median ms | Fixed median ms |
| --- | --- | --- | --- |
| 100 | First | 0.06263–0.06364 | 0.03759–0.03966 |
| 100 | Middle | 0.06077–0.06242 | 0.02620–0.02728 |
| 100 | Last | 0.06005–0.06120 | 0.02624–0.02637 |
| 1,000 | First | 0.50903–0.51099 | 0.17280–0.17620 |
| 1,000 | Middle | 0.49972–0.51555 | 0.17449–0.17640 |
| 1,000 | Last | 0.51642–0.51800 | 0.17857–0.17870 |

The one-procedure control has mixed microsecond-scale changes (last position
0.00243–0.00245 to 0.00262–0.00328 ms). The largest fixed 1,000-procedure batch
average was 0.29478 ms. Cold source stripping and reference classification still
scan the whole module; this measures one warm helper, not the complete refactor
or editor latency.
