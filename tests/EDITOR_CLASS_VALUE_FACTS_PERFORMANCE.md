# Editor class surfaces and diagnostic value facts

Related: https://github.com/WilliamSmithEdward/xlide_vscode/issues/985

Completion and hover built class member surfaces using the same query as
diagnostics. That query scanned class bodies to infer whether object/Variant
members always yield Nothing, Empty or a scalar. Editor completion and hover
consume signatures, types and documentation; the inferred knownValue metadata
is consumed by diagnostic rules.

projectEditorSymbolContextForModule now requests surfaces without these value
facts. Existing ProjectIndex queries include them by default. The two variants
have separate cache keys and member objects, so requesting either first cannot
mutate or strip the other. Source changes invalidate both. No response waits
were added.

## Reproduce

Portable checks verify zero class-value scans in editor projection, complete
diagnostic facts in either cache order, unchanged signatures/documentation,
immutable existing snapshots, and refresh after edits:

~~~powershell
npx vitest run tests/editorClassValueFactsPerformance.test.ts
~~~

The optional benchmark reads the real workbook without writing it. Each fresh
source gets two fresh indexes, then measures class member surfaces with and
without value facts. Timing order alternates; every member is compared after
removing only diagnostic knownValue metadata:

~~~powershell
$env:XLIDE_PERF_WORKBOOK = 'F:\GitHub\xlide\xlide_vscode_testing\ROneCOne_Delegates_Demo.xlsm'
npx vitest run tests/memberFactWorkbookPerformance.test.ts
~~~

For real editor completion and hover, the harness uses a disposable workbook
copy and VS Code 1.139.1:

~~~powershell
npm run compile
npx vscode-test --code-version 1.139.1 --grep 'Actual large class latency'
~~~

## Measurements and remaining work

On the 969,296-character ROneCOne class, matched medians over 21 fresh-source
samples were 11.37 ms with diagnostic value facts and 4.16 ms for editor surfaces.
Maximum samples were 14.29 and 8.68 ms, respectively. This isolates member-surface
construction; parsing and index construction are outside the timed region.

The run of all 17 real-editor checks recorded edit-to-completion-result samples
of 102.81, 71.36, 56.74, 88.06 and 166.04 ms. Warm hover took 6.01 ms; fresh
declaration edit to hover took 71.84 ms. Completion provider traces were 17–65 ms
and semantic coloring still took 75–76 ms. These are individual observations,
not an isolated end-to-end before/after comparison or menu-paint measurements.

Full parsing and semantic/background work remain targets for issue #985.
The optional hover harness can encounter automatic casing between computing and
applying an edit. It now retries rejected edits with current positions, without
sleeping, counts the attempts and includes all attempts in the measured latency.
