# Receiver scope lookup performance

Related: https://github.com/WilliamSmithEdward/xlide_vscode/issues/985

Each receiver chain previously searched all module members to locate its
procedure, then walked the procedure's declarations and module fields to bind
the receiver name. Semantic coloring repeats these queries across many dots.

procedureAtOffset now indexes the immutable AST's procedure spans once and
uses binary lookup. Invalid, unsorted, overlapping or adjacent inclusive spans
retain the old first-match search. Member completion, hover and host semantic
coloring use the same helper for declaration, Set-assignment and With lookup.

Receiver declarations are indexed once per immutable procedure/module AST.
Parameters still precede locals, locals precede module fields, and the existing
first declaration/depth-first traversal wins. Untyped declarations still shadow
host globals. Fresh source edits create new ASTs and therefore new indexes;
WeakMaps follow AST lifetime without retaining older snapshots. No response
waits were added.

## Reproduce

~~~powershell
npx vitest run tests/procedureScopeLookupPerformance.test.ts
~~~

Tests compare procedure lookup against the old first-match search at every
offset in normal and unfinished fixtures. A 1,500-procedure work-count test
verifies no module-member reads on later position queries. Receiver work-count
tests check that local declaration bodies are not walked repeatedly. Additional
cases cover precedence, untyped shadows, duplicate declarations, invalid spans,
source edits and unchanged earlier snapshots.

The optional actual-workbook benchmark reads ROneCOne without modifying it.
AST parsing is outside the timed region; scope/declaration indexes are fresh
for every source snapshot:

~~~powershell
$env:XLIDE_PERF_WORKBOOK = 'F:\GitHub\xlide\xlide_vscode_testing\ROneCOne_Delegates_Demo.xlsm'
npx vitest run tests/receiverScopeWorkbookPerformance.test.ts
~~~

The editor harness copies the workbook into a disposable workspace:

~~~powershell
npm run compile
npx vscode-test --code-version 1.139.1 --grep 'Actual large class latency'
~~~

## Measurements and remaining work

A matched same-process comparison used temporary copies of the preceding
member resolver and host semantic collector, alternated old/new timing order,
and compared every semantic token across 21 fresh ROneCOne snapshots. The old
host-member collector median was 17.05 ms; the indexed version was 4.72 ms.
Maximum samples were 26.24 and 30.22 ms respectively, so this is a median
improvement, not proof that every slow outlier is eliminated. The committed
opt-in benchmark reports current collector timings and checks token stability
across comment-only source edits; it does not depend on temporary baseline files.

All 17 real-editor checks passed. ROneCOne edit-to-completion-result observations
were 97.86, 83.25, 82.66, 65.10 and 61.49 ms. Warm hover was 6.66 ms and fresh
declaration edit to hover was 81.15 ms. Semantic-token provider traces remained
50–64 ms. These command-result measurements are not menu/tooltip paint timings
or an isolated end-to-end before/after comparison.

Full parsing, other semantic collectors and background scheduling still need
work under issue #985.
