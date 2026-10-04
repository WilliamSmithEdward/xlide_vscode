# Reference collection token prefixes

Find References and Rename formerly passed moduleTokens.slice(0, tokenEnd) for each matching name. That copied progressively larger prefixes, even for bare names that the member resolver immediately rejects. The collector now supplies the module's significant tokens through the existing sourceTokens context. The resolver handles its own exact-token boundary check and copies only back to the preceding logical newline, retaining the boundary marker. Its fallback tokenizer handles offsets inside bracketed or suffixed names. This removes the duplicate caller token-end loop and requires no new cache.

## Work counts

Actual collectSymbolReferences with a Library.Greet call on each line, checked complete caller spans and frozen cached token objects/arrays:

| Qualified calls | Baseline copied tokens | Fixed copied tokens | References including declaration |
| ---: | ---: | ---: | ---: |
| 10 | 263 | 40 | 11 |
| 100 | 20,603 | 400 | 101 |
| 1,000 | 2,006,003 | 4,000 | 1,001 |

Two regressions exercise both includeDeclaration settings. They count actual cached-token array slices and assert the complete ordered reference result. Both fail the work bound on df1bc7ed and pass with the fix.

## Timings

Windows, Ryzen 7 9800X3D, Node v24.18.0; baseline df1bc7ed. Four isolated processes ran baseline/fixed/fixed/baseline, with three warmups and fifteen measured rounds per configuration. Each table range spans the two processes for that version, in milliseconds; p95 is the largest of fifteen measured rounds. All sixteen complete-result hashes match in all four processes.

The benchmark times the complete actual collector with a warm project and source caches. Project construction, source preparation and complete-output assertions are outside the timer. Qualified calls use Library.Greet; bare calls use Greet; continued calls use Library. followed by underscore continuation and Greet; With calls use .Greet inside With Library. Each module also contains the expected declaration.

| Calls | Style | Baseline median ms | Fixed median ms | Baseline p95 ms | Fixed p95 ms |
| ---: | --- | ---: | ---: | ---: | ---: |
| 1 | qualified | 0.034–0.044 | 0.035–0.037 | 0.092–0.097 | 0.099–0.107 |
| 1 | bare | 0.024–0.024 | 0.022–0.024 | 0.035–0.071 | 0.032–0.033 |
| 1 | continued | 0.020–0.020 | 0.021–0.022 | 0.040–0.040 | 0.046–0.058 |
| 1 | with | 0.041–0.044 | 0.029–0.029 | 0.057–0.060 | 0.034–0.042 |
| 10 | qualified | 0.056–0.056 | 0.057–0.057 | 0.064–0.067 | 0.073–0.088 |
| 10 | bare | 0.041–0.046 | 0.039–0.042 | 0.169–0.170 | 0.170–0.175 |
| 10 | continued | 0.052–0.054 | 0.055–0.056 | 0.069–0.198 | 0.144–0.161 |
| 10 | with | 0.167–0.186 | 0.085–0.096 | 1.403–1.921 | 0.775–1.034 |
| 100 | qualified | 0.179–0.195 | 0.232–0.245 | 0.273–0.335 | 1.870–1.942 |
| 100 | bare | 0.164–0.167 | 0.121–0.124 | 0.220–0.233 | 0.270–0.327 |
| 100 | continued | 0.187–0.197 | 0.236–0.242 | 0.494–0.525 | 0.264–0.308 |
| 100 | with | 2.353–2.522 | 0.532–0.550 | 2.825–2.849 | 0.869–0.872 |
| 1000 | qualified | 3.148–3.255 | 2.679–2.764 | 4.335–4.706 | 3.267–3.499 |
| 1000 | bare | 0.903–1.013 | 0.813–0.881 | 1.354–1.549 | 1.039–1.097 |
| 1000 | continued | 2.998–3.021 | 2.704–2.727 | 3.425–3.651 | 3.021–3.545 |
| 1000 | with | 227.416–228.507 | 32.977–33.534 | 230.637–236.969 | 34.035–34.991 |

At 1,000 qualified calls, the complete collector improves from 3.148–3.255 ms to 2.679–2.764 ms; With calls improve from 227.416–228.507 ms to 32.977–33.534 ms. The binary search and bounded-prefix path have a cost at some smaller sizes: 100 qualified calls increase from 0.179–0.195 ms to 0.232–0.245 ms, and 100 continued calls increase from 0.187–0.197 ms to 0.236–0.242 ms. Single continued-call controls also increase slightly. These data establish reduced copying and the measured warm-workload changes, not universally faster reference collection.

No cold whole-project construction, Office execution or VS Code UI latency was measured. The residual With cost is a separate audit candidate.

Reproduce from the repository root:

```powershell
node scripts/benchmark-reference-prefixes.mjs --baseline=df1bc7ed
node scripts/benchmark-reference-prefixes.mjs
node scripts/benchmark-reference-prefixes.mjs
node scripts/benchmark-reference-prefixes.mjs --baseline=df1bc7ed
npm exec vitest run tests/referenceCollectorPrefixWork.test.ts tests/vbaReferenceResolution.test.ts tests/vbaBangAndBracketReferences.test.ts tests/vbaReferenceKinds.test.ts tests/vbaProjectReferences.test.ts tests/vbaFormReferenceWiring.test.ts
npm run check-types
npm test -- --run
```

## Validation

The six focused files pass 54 tests, including the two new work regressions. Type checking passes. The full suite passes 651 files and 13,220 tests with 13 skipped on df1bc7ed plus this change.

An additional ad hoc baseline/fixed differential audit compares complete reference results across 360 fixtures and 2,880 invocations. It covers qualified/bare/With calls, nested With, underscores, comments/strings, colons, inline If, bracketed and suffixed names, dangling dots, local/module/procedure shadows, unrelated class receivers, attributes, LF/CRLF/CR endings, invocation from either declaration or caller, and both declaration settings. Complete references, ambiguity and kinds match with full-source cached token objects/arrays frozen and zero uncaught exceptions. This does not claim frozen ASTs or exhaustive repository coverage.
