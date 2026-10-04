# Assignment significant-token reuse

The assignment rules copied arrays to remove comments from input that already contains only significant tokens. statementTokensCached excludes comments/newlines in both the derived-token path and its fallback lexer. Bare Set and array-element target extractors slice those tokens, and Mid's argument slice originates from statementTokensAfterLeadingLabel over the same cache.

Remove the array-target filter, three Set RHS filters and the Mid argument filter. One Set RHS alias now serves control, literal, held-object and compatibility checks. Existing slices required to select targets/arguments remain. These paths only read tokens; no cached arrays or token objects are changed. Other filters whose input contracts were not established by this audit remain unchanged.

## Deterministic cost

Scratch-only instrumentation counts the five removed filter sites with 100 writes; each filter increments its allocation counter and each predicate visit increments its token-check counter:

| Mode | Before arrays / token checks | After arrays / token checks | Diagnostics |
| --- | ---: | ---: | ---: |
| Object / Nothing | 400 / 700 | 0 / 0 | 0 |
| Long / Nothing | 300 / 600 | 0 / 0 | 100 |
| Variant / Nothing | 300 / 600 | 0 / 0 | 0 |
| Object / Unknown with 100 arguments | 400 / 81,100 | 0 / 0 | 0 |
| Mid literal target | 100 / 300 | 0 / 0 | 100 |
| Mid variable target | 100 / 300 | 0 / 0 | 0 |

These are real filter calls on the rule paths, not simulated costs. Zero means these five sites are removed, not that the rule allocates no other arrays. Counters are absent from production and timing runs.

## Reproduction and timings

From the repository root with installed dependencies:

```powershell
node scripts/benchmark-assignment-significant-tokens.mjs --baseline=bdc09289
node scripts/benchmark-assignment-significant-tokens.mjs
```

The baseline replaces assignments.ts and typeInference.ts from bdc09289, after the interface-sharing fix, in the current harness; only assignments.ts changes in this PR. Fixtures contain the listed number of writes. Set modes vary declared Object/Long/Variant types or use a 100-argument Unknown call as the RHS. Mid modes assign through a literal or variable target. No-set uses ordinary c=1 statements while visiting the Set rule, exercising its array-target rejection path. Each sample checks one diagnostic per write for scalar Set and Mid literals, and zero for other modes.

Set fixtures time checkSetAssignments and its statement visitors; Mid fixtures time checkMidStatementLiteralTarget. Parsing, tokenization and binding are outside timing. Parsed bodies/tokens are reused, with fresh bound roots and completion caches and a root identity assertion for each sample. These are rule timings with reused parsed bodies, not cold dataflow or full analyzer/editor latency. Three warmups precede 15 samples. Before/after runs were sequential after tests and differential probes completed, with no concurrent audit workloads.

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor           ; measured locally on 2026-10-03. Times in milliseconds.

| Writes / mode | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: |
| 100-set-object | 0.294 | 0.340 | 0.511 | 0.922 |
| 100-set-scalar | 0.094 | 0.086 | 0.832 | 1.472 |
| 100-set-variant | 0.097 | 0.072 | 0.164 | 0.138 |
| 100-set-long-rhs | 1.191 | 0.614 | 2.841 | 1.241 |
| 100-mid-literal | 0.044 | 0.032 | 0.254 | 0.063 |
| 100-mid-variable | 0.024 | 0.018 | 0.043 | 0.035 |
| 100-no-set | 0.019 | 0.015 | 0.243 | 0.027 |
| 1000-set-object | 1.415 | 1.625 | 2.554 | 2.687 |
| 1000-set-scalar | 0.418 | 0.317 | 1.403 | 0.383 |
| 1000-set-variant | 0.472 | 0.320 | 0.981 | 0.371 |
| 1000-set-long-rhs | 7.849 | 5.396 | 13.323 | 6.736 |
| 1000-mid-literal | 0.251 | 0.192 | 0.288 | 0.208 |
| 1000-mid-variable | 0.181 | 0.160 | 0.218 | 0.180 |
| 1000-no-set | 0.073 | 0.062 | 0.147 | 0.085 |
| 3000-set-object | 3.551 | 3.102 | 5.068 | 4.328 |
| 3000-set-scalar | 1.007 | 0.997 | 1.667 | 1.418 |
| 3000-set-variant | 1.054 | 0.971 | 1.812 | 1.777 |
| 3000-set-long-rhs | 24.136 | 15.031 | 25.354 | 16.112 |
| 3000-mid-literal | 0.589 | 0.524 | 0.671 | 1.037 |
| 3000-mid-variable | 0.544 | 0.481 | 1.181 | 0.573 |
| 3000-no-set | 0.227 | 0.180 | 0.936 | 0.221 |

At 3,000 long-RHS Set writes the median improves from 24.136 to 15.031 ms, about 38 percent. Short expressions and Mid cases are smaller and mixed: Object medians/p95 regress at 100/1,000 writes, scalar p95 regresses at 100 writes, and 3,000-write Mid literal p95 rises from 0.671 to 1.037 ms despite a lower median. The deterministic redundant copies disappear, but these timings do not establish a universal latency improvement. Other inference, target extraction, traversal and allocation costs remain.

## Validation

- Type check passed.
- 213 focused tests passed, including 13 new comment/label/continuation and cache-immutability regressions; the final cleaned test scaffolding was rechecked with all 13 passing.
- Full suite: 626 files, 12,856 tests passed, 13 skipped.
- New regressions freeze actual cached token arrays and objects before the rules run, covering trailing comments, named/numeric labels, single-line branches, continuations, array targets, unknown expressions, Mid/Mid$/MidB literal spans with quoted apostrophes, declared/ReDim intrinsic shadows and scalar/Variant literal diagnostics.
- 3,000 differential cases preserve complete analyzer and direct Set/Mid outputs with no internal errors. Fixtures include comments/labels/continuations and Mid literal/variable/shadow cases, metadata/compatibility controls, conditional declarations, arrays, branches/loops and four host models.

Generated counting and differential probes stay in audit scratch material. The timing benchmark and regression tests are committed.
