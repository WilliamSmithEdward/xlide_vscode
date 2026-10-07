# Reference collection With indexes

The member resolver already has a per-procedure With index cache, but Find References and Rename did not supply it. Each leading-dot member rebuilt the complete procedure's index. The reference collector now gives each module's collection pass a fresh withScanCache Map. The resolver fills it lazily, once per procedure. The Map is discarded after that module's pass, so later collections, edited source and other modules cannot inherit those facts. No global cache or new index implementation is introduced.

Module-level malformed source retains the existing uncached fallback because its window has no procedure key. Nested With expressions, labels and colon boundaries remain handled by the existing index.

## Deterministic work

Actual public collectSymbolReferences on a single With block with frozen real cached token objects/arrays and exact complete ordered result assertions:

| Leading-dot references | Baseline index builds | Fixed index builds | Baseline tokens scanned by builds | Fixed tokens scanned by builds |
| ---: | ---: | ---: | ---: | ---: |
| 10 | 10 | 1 | 430 | 43 |
| 100 | 100 | 1 | 31,300 | 313 |
| 1,000 | 1,000 | 1 | 3,013,000 | 3,013 |

Three new work regressions observe stable end getters on actual frozen cached tokens for LF, CRLF and CR. All fail the read bound on c3fb6da5 while their complete outputs already match expectations. Two controls cover separate procedures and changed receivers across collections. All five pass with the fix.

## Timings

Windows, Ryzen 7 9800X3D, Node v24.18.0. Baseline c3fb6da5 includes the bounded-prefix fix (#909). Four isolated processes ran baseline/fixed/fixed/baseline with three warmups and fifteen measured rounds. The ranges span the two processes for each version, in milliseconds; p95 is the largest of fifteen samples. All twenty complete output hashes match across all four processes.

Timings cover the complete actual collector with warm project and source caches. Source/project construction and complete-output assertions are outside the timer. Each collection gets a fresh With Map, so single-reference controls include its first index build even though parser/lexer caches are warm.

Qualified, bare and continued calls are non-With controls. The with fixture uses the resolver's Library module receiver; typed-with uses a Library class, Dim receiver As Library, Set receiver = New Library and With receiver. It provides an ordinary object-receiver fixture alongside the module-receiver workload. No Office execution, cold whole-project construction or VS Code UI latency was measured.

| Calls | Style | Baseline median ms | Fixed median ms | Baseline p95 ms | Fixed p95 ms |
| ---: | --- | ---: | ---: | ---: | ---: |
| 1 | qualified | 0.036–0.043 | 0.036–0.041 | 0.096–0.097 | 0.100–0.137 |
| 1 | bare | 0.024–0.024 | 0.021–0.024 | 0.030–0.031 | 0.030–0.036 |
| 1 | continued | 0.022–0.025 | 0.021–0.024 | 0.046–0.081 | 0.046–0.146 |
| 1 | with | 0.029–0.031 | 0.030–0.032 | 0.036–0.042 | 0.041–0.156 |
| 1 | typed-with | 0.024–0.025 | 0.025–0.027 | 0.033–0.038 | 0.037–0.040 |
| 10 | qualified | 0.057–0.061 | 0.058–0.059 | 0.192–0.202 | 0.202–0.203 |
| 10 | bare | 0.041–0.043 | 0.042–0.042 | 0.061–0.158 | 0.050–0.106 |
| 10 | continued | 0.048–0.054 | 0.054–0.055 | 0.072–0.238 | 0.084–0.211 |
| 10 | with | 0.086–0.109 | 0.068–0.072 | 0.778–1.020 | 0.747–1.005 |
| 10 | typed-with | 0.060–0.063 | 0.048–0.051 | 1.419–1.468 | 0.077–0.164 |
| 100 | qualified | 0.183–0.186 | 0.193–0.209 | 0.420–0.474 | 0.460–0.486 |
| 100 | bare | 0.114–0.126 | 0.119–0.122 | 0.363–0.363 | 0.267–0.363 |
| 100 | continued | 0.248–0.252 | 0.255–0.256 | 0.421–0.508 | 0.363–0.546 |
| 100 | with | 0.500–0.517 | 0.184–0.185 | 0.788–0.861 | 0.231–0.240 |
| 100 | typed-with | 0.455–0.469 | 0.160–0.161 | 0.605–0.624 | 0.374–0.594 |
| 1000 | qualified | 2.846–2.897 | 2.901–2.917 | 3.370–3.905 | 3.301–3.319 |
| 1000 | bare | 0.832–0.845 | 0.793–0.835 | 1.164–1.785 | 1.171–1.383 |
| 1000 | continued | 2.817–2.847 | 2.811–2.864 | 3.185–3.470 | 2.926–3.295 |
| 1000 | with | 34.009–34.404 | 3.014–3.109 | 36.236–37.978 | 3.556–3.842 |
| 1000 | typed-with | 34.297–34.600 | 1.242–1.303 | 42.929–43.412 | 1.544–1.891 |

At 1,000 typed class-instance With references, complete collection improves from 34.297–34.600 ms to 1.242–1.303 ms. The module-receiver With fixture improves from 34.009–34.404 ms to 3.014–3.109 ms. Non-With controls are mixed: 100 qualified calls increase from 0.183–0.186 ms to 0.193–0.209 ms, and 1,000 qualified calls increase from 2.846–2.897 ms to 2.901–2.917 ms. Single-reference With controls have a small cost as well. These measurements do not establish a universal whole-analyzer improvement.

Reproduce from the repository root:

```powershell
node scripts/benchmark-reference-prefixes.mjs --baseline=c3fb6da5
node scripts/benchmark-reference-prefixes.mjs
node scripts/benchmark-reference-prefixes.mjs
node scripts/benchmark-reference-prefixes.mjs --baseline=c3fb6da5
npm exec vitest run tests/referenceWithScanWork.test.ts tests/referenceCollectorPrefixWork.test.ts tests/vbaReferenceResolution.test.ts tests/vbaBangAndBracketReferences.test.ts tests/vbaReferenceKinds.test.ts tests/vbaProjectReferences.test.ts tests/vbaFormReferenceWiring.test.ts
npm run check-types
npm test -- --run
```

## Behavior validation

Seven focused files pass 59 tests; type checking passes. After rebasing onto main945ac8de (including #907 and #909), type checking and the final full suite pass: 653 files and 13,242 tests with 13 skipped. Runtime and tests were unchanged during the final validation; the only subsequent edits update this document.

An ad hoc baseline/fixed differential check compares complete reference results, ambiguity and kinds for 450 fixtures and 3,600 invocations. It covers nested and unclosed With, module-level malformed With, stray End With, colon/label boundaries, continuation, comments/strings, inline If, bracketed/suffixed names, local/module/procedure shadowing, unrelated class receivers, attributes, three newline forms, declaration/caller invocation sites and both declaration settings. All outputs match with cached token objects/arrays frozen and no uncaught exceptions. This is token-ownership evidence, not frozen AST or exhaustive repository coverage.

The final main-based differential run repeats all 3,600 comparisons against 945ac8de and matches complete outputs with frozen cached tokens and zero exceptions. Timings above retain their explicitly named c3fb6da5 baseline.
