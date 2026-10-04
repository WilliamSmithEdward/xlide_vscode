# Statement token prefix reuse

Related: https://github.com/WilliamSmithEdward/xlide_vscode/issues/985

Semantic coloring and analysis repeatedly request tokens for a module's
statements. The two-entry cache previously rebuilt those arrays for every new
source string. It now borrows immutable arrays wholly before the changed
physical line for small edits in large modules. Span maps are copied rather
than retaining previous cache entries, so typing cannot create an unbounded
chain of source snapshots. Changed-line and suffix spans are derived normally.
No response waits were added.

Reuse requires at least 4,096 characters, a recent snapshot within 128
characters in length, and at least 4,096 unchanged prefix characters. Only
finite, nonnegative, ordered spans ending before the changed line qualify.
Cold, early, large and unrelated edits keep the original derivation path.

## Verification

~~~powershell
npx vitest run tests/statementTokenPrefixPerformance.test.ts
$env:XLIDE_PERF_WORKBOOK = 'F:\GitHub\xlide\xlide_vscode_testing\ROneCOne_Delegates_Demo.xlsm'
npx vitest run tests/statementTokenWorkbookPerformance.test.ts
npm run compile
npx vscode-test --code-version 1.139.1 --grep 'Actual large class latency'
~~~

Portable tests verify that 200 unchanged statement arrays are reused with no
module-token queries, while edited lines are rebuilt. They compare every tail
insertion/deletion against cold derivation, including Unicode, CRLF,
continuations, colon-separated statements, trivia and token positions. Earlier
snapshots remain unchanged. The read-only workbook probe checks every statement
token field against a cold cache for 21 new snapshots; it has no temporary
baseline dependency. The editor harness edits a disposable workbook copy.

## Measurements and limits

The extension-host profile attributed 25.5 ms across the typing sequence to
statement-token derivation. A temporary preceding implementation enabled a
matched same-process probe on ROneCOne's 9,915 statement spans. Both versions
were warmed on the original snapshot, module lexing was outside the timed
region, order alternated, and every resulting token array matched across 21
fresh snapshots. Median construction fell from 10.42 to 3.22 ms; maximum
samples were 15.13 and 5.58 ms. A separate current-only cold-parity run had a
2.37 ms median and a 14.39 ms maximum, so slow outliers remain.

All 17 editor checks passed. Actual large-class edit-to-completion-result
samples were 43.74, 46.81, 61.86, 96.88 and 37.18 ms. Warm hover took 24.14 ms
and fresh-edit hover 30.53 ms. Semantic-token provider traces still reached
57 ms and full diagnostic traces 715 ms. These are command-result observations,
not menu/tooltip paint measurements or a matched end-to-end comparison.
Completion remains variable despite the component improvement; issue #985
stays open for other collectors, cold/early edits and background scheduling.
