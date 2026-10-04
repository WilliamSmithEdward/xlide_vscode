# Parser prefix reuse performance

Related: https://github.com/WilliamSmithEdward/xlide_vscode/issues/985

Small edits in large modules previously reparsed every procedure. The bounded
module cache now shares complete, unchanged members before a safe logical-line
boundary and parses the remaining tokens with their original absolute offsets.
Earlier AST snapshots remain immutable. No response waits were added.

Reuse requires a source of at least 4,096 characters and a cached snapshot
within 128 characters in length. Any previous conditional directive prevents
reuse; diagnostics limit reuse to the prefix before the first diagnostic.
Unclosed members and partially changed logical lines are not reused. Cold
snapshots, early edits, large changes and unsupported boundaries retain full
parsing. The cache still holds at most eight snapshots.

## Verification

`npx vitest run tests/parserPrefixReusePerformance.test.ts` compares incremental
results with cold full parses after insertions and deletions at every offset of
a tail fixture. It checks all AST fields, diagnostics, spans, immutable earlier
snapshots, member sharing, conditional directives, error recovery, CRLF edits,
colon-separated declarations and extended closing keywords.

The optional read-only workbook probe measures early, middle and late edits,
warms lexing outside the timed region, and verifies every result against a cold
full parse:

~~~powershell
$env:XLIDE_PERF_WORKBOOK = 'F:\GitHub\xlide\xlide_vscode_testing\ROneCOne_Delegates_Demo.xlsm'
npx vitest run tests/parserPrefixWorkbookPerformance.test.ts
npm run compile
npx vscode-test --code-version 1.139.1 --grep 'Actual large class latency'
~~~

The editor harness edits a disposable workbook copy.

## Observations

A temporary copy of the preceding full parser enabled a matched same-process
comparison on the actual 969,296-character ROneCOne class. Fifteen source edits
per position alternated timing order, with lexing warmed outside both timed
regions. All 45 complete ASTs and diagnostics matched.

| Edit position | Full parser median | Prefix reuse median |
| --- | ---: | ---: |
| 1% | 64.98 ms | 59.94 ms |
| 50% | 33.28 ms | 19.43 ms |
| 99% | 31.26 ms | 1.54 ms |

These position-specific component measurements include machine/warmup variation;
they do not show equivalent benefits for early edits. The committed probe does
not depend on temporary baseline files.

All 17 editor surface checks passed. Actual ROneCOne edit-to-completion-result
samples were 39.15, 46.91, 39.19, 37.79 and 44.01 ms. Warm hover took 30.94 ms;
fresh declaration edit to hover took 48.16 ms. The preceding run observed
61–98 ms completion and 81 ms fresh hover, but these separate editor runs are
not an isolated matched before/after comparison. Commands measure returned
results, not completion-menu or mouse-tooltip painting. Background diagnostic
traces still reached 760 ms, and semantic coloring reached 57 ms. Issue #985
remains open for cold/early edits, other collectors and background scheduling.
