# Assignment-target parenthesis classification

Reference write classification previously scanned the same nested parenthesis suffix for every name on an assignment target, including names whose offsets were not requested. The fix skips unrequested target names and builds a local parenthesis endpoint map only after a second suffix query. The first suffix retains the existing scan. There is no persistent cache or metadata lifetime change.

For malformed source, an opening parenthesis can close after the detected assignment operator. The endpoint scan continues until the outstanding stack is empty, preserving the previous full-segment fallback; unmatched openings still advance one token. Set, Let, LSet and RSet share the same target routine.

## Deterministic work

The actual classifyReferenceKinds API consumed real tokenizeCached tokens with counted, stable rawText getters and frozen token objects/arrays. Only Value was requested in a(f(f(...1...))).Value = 1; its result remained write.

| Nested calls | Tokens | Baseline rawText reads | Fixed rawText reads |
| ---: | ---: | ---: | ---: |
| 5 | 23 | 141 | 24 |
| 100 | 308 | 26,266 | 309 |
| 1,000 | 3,008 | 2,512,516 | 3,009 |

Regression tests also bound all-name queries for each of the five assignment prefixes, so skipping irrelevant names cannot hide the remaining overlapping scans.

## Timings

Windows, Ryzen 7 9800X3D, Node v24.18.0; baseline df1bc7ed. Four isolated processes ran in baseline/fixed/fixed/baseline order. Each configuration had three warmups and fifteen measured rounds. The table reports the range across the two processes for each version, in milliseconds. p95 is the largest of fifteen measured samples. Values below 0.001 ms round to 0.000; these tiny observations are not reliable speedup estimates.

The benchmark times the actual classifier. Source/offset preparation and complete ordered-map assertions are outside the timer. Cached-token runs reuse source; fresh-token runs vary a trailing comment to include lexing. Terminal queries request only the assigned root/member, while all queries include every identifier. Depth zero uses a(i) = 1; other depths use the nested-call member target above. All 32 complete result hashes matched in all four processes.

| Depth | Statements | Requested names | Tokens | Baseline median ms | Fixed median ms | Baseline p95 ms | Fixed p95 ms |
| ---: | ---: | --- | --- | ---: | ---: | ---: | ---: |
| 0 | 1 | terminal | cached-tokens | 0.003–0.003 | 0.003–0.004 | 0.013–0.013 | 0.013–0.016 |
| 0 | 1 | terminal | fresh-tokens | 0.006–0.006 | 0.005–0.006 | 0.013–0.013 | 0.012–0.013 |
| 0 | 1 | all | cached-tokens | 0.001–0.001 | 0.001–0.001 | 0.001–0.001 | 0.002–0.002 |
| 0 | 1 | all | fresh-tokens | 0.005–0.005 | 0.005–0.005 | 0.018–0.029 | 0.015–0.016 |
| 0 | 100 | terminal | cached-tokens | 0.042–0.048 | 0.040–0.043 | 0.103–0.124 | 0.273–0.335 |
| 0 | 100 | terminal | fresh-tokens | 0.101–0.107 | 0.091–0.104 | 0.472–0.536 | 0.358–0.452 |
| 0 | 100 | all | cached-tokens | 0.026–0.026 | 0.026–0.026 | 0.048–0.050 | 0.038–0.038 |
| 0 | 100 | all | fresh-tokens | 0.072–0.073 | 0.070–0.071 | 0.082–0.085 | 0.202–0.216 |
| 5 | 1 | terminal | cached-tokens | 0.001–0.001 | 0.001–0.001 | 0.001–0.001 | 0.001–0.001 |
| 5 | 1 | terminal | fresh-tokens | 0.007–0.007 | 0.007–0.007 | 0.105–0.167 | 0.119–0.195 |
| 5 | 1 | all | cached-tokens | 0.001–0.001 | 0.002–0.002 | 0.001–0.005 | 0.002–0.012 |
| 5 | 1 | all | fresh-tokens | 0.007–0.007 | 0.008–0.008 | 0.016–0.019 | 0.019–0.020 |
| 5 | 100 | terminal | cached-tokens | 0.068–0.074 | 0.072–0.078 | 0.191–0.306 | 0.215–0.289 |
| 5 | 100 | terminal | fresh-tokens | 0.355–0.364 | 0.360–0.366 | 0.764–0.878 | 0.686–1.119 |
| 5 | 100 | all | cached-tokens | 0.045–0.049 | 0.063–0.064 | 0.056–0.071 | 0.598–0.626 |
| 5 | 100 | all | fresh-tokens | 0.293–0.313 | 0.320–0.340 | 0.860–1.096 | 0.713–0.784 |
| 100 | 1 | terminal | cached-tokens | 0.019–0.019 | 0.005–0.005 | 0.019–0.022 | 0.005–0.006 |
| 100 | 1 | terminal | fresh-tokens | 0.033–0.033 | 0.019–0.019 | 0.033–0.038 | 0.019–0.019 |
| 100 | 1 | all | cached-tokens | 0.019–0.019 | 0.008–0.008 | 0.019–0.020 | 0.008–0.008 |
| 100 | 1 | all | fresh-tokens | 0.033–0.034 | 0.034–0.036 | 0.037–0.054 | 0.048–0.049 |
| 100 | 100 | terminal | cached-tokens | 2.052–2.102 | 0.773–0.779 | 2.127–3.019 | 1.552–1.627 |
| 100 | 100 | terminal | fresh-tokens | 4.803–4.916 | 3.550–3.607 | 6.649–6.967 | 5.167–5.698 |
| 100 | 100 | all | cached-tokens | 1.780–1.846 | 1.450–1.453 | 1.844–5.016 | 4.418–4.890 |
| 100 | 100 | all | fresh-tokens | 3.264–3.264 | 2.847–2.898 | 7.987–8.233 | 6.719–7.378 |
| 1000 | 1 | terminal | cached-tokens | 1.034–1.069 | 0.030–0.030 | 1.070–1.100 | 0.040–0.040 |
| 1000 | 1 | terminal | fresh-tokens | 1.206–1.208 | 0.170–0.174 | 1.268–1.430 | 2.110–2.353 |
| 1000 | 1 | all | cached-tokens | 1.048–1.059 | 0.081–0.081 | 1.091–1.098 | 0.109–0.114 |
| 1000 | 1 | all | fresh-tokens | 1.196–1.205 | 0.218–0.218 | 3.249–3.256 | 0.239–0.245 |

Deep targets improve substantially: a depth-1,000 single target with all names requested drops from 1.048–1.059 ms to 0.081 ms with cached tokens. Shallow all-name batches have a cost: depth five, 100 statements increases from 0.045–0.049 ms to 0.063–0.064 ms cached, and from 0.293–0.313 ms to 0.320–0.340 ms with fresh tokens. A depth-100 single all-name fresh-token query also slightly increases. Simple one-suffix controls are similar. These are classifier measurements, not a universal whole-analyzer improvement.

These nested targets are syntactic analyzer stress fixtures. They do not establish that Office compiles arbitrarily deep or long physical lines, and no Office or VS Code UI timing was measured.

Reproduce from the repository root:

```powershell
node scripts/benchmark-reference-target-parens.mjs --baseline=df1bc7ed
node scripts/benchmark-reference-target-parens.mjs
node scripts/benchmark-reference-target-parens.mjs
node scripts/benchmark-reference-target-parens.mjs --baseline=df1bc7ed
npm exec vitest run tests/referenceTargetParenWork.test.ts tests/vbaReferenceKinds.test.ts
npm run check-types
npm test -- --run
```

## Behavior validation

Seventeen new regression/control tests passed; ten deterministic work tests failed on the baseline. The focused reference suites passed 27 tests. A separate differential audit compared complete ordered Maps (including default reads, duplicates and unusual offsets) for 8,256 corpus sources in four offset modes: 33,024 comparisons. Another 5,832 generated comparisons covered six statement prefixes, nested/adjacent/malformed parentheses and LF/CRLF/CR endings. All matched, with cached token objects and arrays frozen and zero uncaught exceptions. This is token immutability evidence, not an AST-freezing claim.

Type checking passed. The final full suite passed 651 files and 13,235 tests, with 13 skipped. Validation used df1bc7ed plus this change; the only subsequent source edit moved a documentation comment.
