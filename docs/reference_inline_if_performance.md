# Inline If reference classification

The syntactic reference classifier formerly rescanned and copied the full remaining inline If tail at every nested head. The outer classifier already splits every parenthesis-depth-zero Else into ordered fragments. The fix peels nested If/ElseIf heads within those fragments and sends only the final leaf to the existing statement classifier. A shared bounded Then scanner removes duplicated scanning code.

After a depth-zero Then, the parenthesis depth equals the fragment's starting depth. Its tail therefore cannot contain another depth-zero Else: the caller already split every such Else. Malformed conditions with no depth-zero Then retain the existing default-read behavior. The original outer Else partition and leaf assignment/declaration/readwrite rules remain the source of classification and ordered writes. No new persistent cache or separate statement parser is introduced.

## Deterministic work

Actual classifyReferenceKinds on If c Then repeated before x = 1, querying x alone, with frozen real tokenizeCached token objects/arrays and stable rawText/start getters. Complete ordered Maps assert x remains write.

| Nested heads | Tokens | Baseline rawText reads | Fixed rawText reads | Baseline start reads | Fixed start reads |
| ---: | ---: | ---: | ---: | ---: | ---: |
| 10 | 33 | 274 | 77 | 178 | 34 |
| 100 | 303 | 20,704 | 707 | 15,253 | 304 |
| 1,000 | 3,003 | 2,007,004 | 7,007 | 1,502,503 | 3,004 |
| 5,000 | 15,003 | 50,035,004 | 35,007 | 37,512,503 | 15,004 |

Four baseline-failing work tests also cover both If and ElseIf heads with sparse and all-name queries. Eight controls preserve nested/Else fragment order, missing Then, malformed parentheses and declaration/Let targets. All twelve pass with the fix. No stack exception occurred in the measured baseline cases; this finding is repeated work, not a claimed reproduced stack crash.

## Timings

Windows, Ryzen 7 9800X3D, Node v24.18.0; baseline 0f685d45. Four isolated processes ran baseline/fixed/fixed/baseline. Each configuration has three warmups and fifteen measured rounds. Ranges span the two processes per version, in milliseconds; p95 is the largest of fifteen samples. All 32 complete ordered-map result hashes match in all four processes.

The existing classifier benchmark accepts --shape=inline-if instead of adding a duplicate benchmark. Its default parenthesis-target mode also completed a smoke run with all 32 output checks. Timed calls use the actual classifier. Source/offset preparation and complete output assertions are outside timing. Cached-token mode reuses source; fresh-token mode varies a trailing comment to include lexing. Depth zero is the unchanged a(i) = 1 control; positive depths use repeated If c Then heads before x = 1. Terminal queries request only the write target; all queries request every identifier.

| Depth | Statements | Requested names | Tokens | Baseline median ms | Fixed median ms | Baseline p95 ms | Fixed p95 ms |
| ---: | ---: | --- | --- | ---: | ---: | ---: | ---: |
| 0 | 1 | terminal | cached-tokens | 0.003–0.003 | 0.003–0.003 | 0.012–0.013 | 0.012–0.013 |
| 0 | 1 | terminal | fresh-tokens | 0.005–0.005 | 0.005–0.005 | 0.013–0.014 | 0.013–0.014 |
| 0 | 1 | all | cached-tokens | 0.001–0.001 | 0.001–0.001 | 0.001–0.001 | 0.001–0.001 |
| 0 | 1 | all | fresh-tokens | 0.005–0.005 | 0.005–0.005 | 0.016–0.017 | 0.016–0.016 |
| 0 | 100 | terminal | cached-tokens | 0.047–0.047 | 0.047–0.054 | 0.159–0.286 | 0.315–0.317 |
| 0 | 100 | terminal | fresh-tokens | 0.093–0.114 | 0.099–0.111 | 0.338–0.577 | 0.362–0.452 |
| 0 | 100 | all | cached-tokens | 0.028–0.029 | 0.026–0.028 | 0.044–0.049 | 0.046–0.077 |
| 0 | 100 | all | fresh-tokens | 0.074–0.074 | 0.071–0.072 | 0.141–0.273 | 0.213–0.236 |
| 5 | 1 | terminal | cached-tokens | 0.005–0.005 | 0.003–0.003 | 0.060–0.182 | 0.004–0.004 |
| 5 | 1 | terminal | fresh-tokens | 0.011–0.011 | 0.011–0.016 | 0.020–0.021 | 0.127–0.135 |
| 5 | 1 | all | cached-tokens | 0.004–0.004 | 0.002–0.003 | 0.008–0.013 | 0.003–0.004 |
| 5 | 1 | all | fresh-tokens | 0.010–0.010 | 0.008–0.009 | 0.021–0.023 | 0.023–0.027 |
| 5 | 100 | terminal | cached-tokens | 0.212–0.216 | 0.084–0.089 | 0.419–0.472 | 0.388–0.480 |
| 5 | 100 | terminal | fresh-tokens | 0.471–0.506 | 0.474–0.496 | 1.110–1.145 | 1.267–1.386 |
| 5 | 100 | all | cached-tokens | 0.152–0.158 | 0.120–0.120 | 0.958–1.177 | 0.154–0.295 |
| 5 | 100 | all | fresh-tokens | 0.243–0.244 | 0.203–0.208 | 1.142–1.262 | 1.191–1.219 |
| 100 | 1 | terminal | cached-tokens | 0.216–0.216 | 0.006–0.006 | 0.220–0.224 | 0.007–0.008 |
| 100 | 1 | terminal | fresh-tokens | 0.256–0.258 | 0.029–0.030 | 0.410–0.436 | 0.031–0.039 |
| 100 | 1 | all | cached-tokens | 0.120–0.121 | 0.008–0.008 | 0.123–0.123 | 0.008–0.009 |
| 100 | 1 | all | fresh-tokens | 0.229–0.232 | 0.031–0.032 | 0.240–0.428 | 0.057–0.063 |
| 100 | 100 | terminal | cached-tokens | 16.887–16.921 | 0.983–1.070 | 17.421–18.005 | 2.020–2.275 |
| 100 | 100 | terminal | fresh-tokens | 22.184–22.299 | 6.376–6.512 | 24.609–33.995 | 12.193–12.799 |
| 100 | 100 | all | cached-tokens | 12.699–12.888 | 1.135–1.289 | 12.902–17.650 | 1.297–4.025 |
| 100 | 100 | all | fresh-tokens | 17.836–17.890 | 8.242–8.250 | 25.139–25.910 | 9.313–9.537 |
| 1000 | 1 | terminal | cached-tokens | 14.892–14.906 | 0.063–0.065 | 15.090–15.598 | 0.080–0.123 |
| 1000 | 1 | terminal | fresh-tokens | 16.731–16.805 | 0.301–0.313 | 17.175–17.309 | 2.860–3.068 |
| 1000 | 1 | all | cached-tokens | 11.300–11.332 | 0.078–0.088 | 11.597–11.607 | 0.083–0.094 |
| 1000 | 1 | all | fresh-tokens | 12.587–12.607 | 0.316–0.324 | 13.072–13.127 | 1.349–1.576 |

A depth-1,000 terminal query improves from 14.892–14.906 ms to 0.063–0.065 ms cached, and from 16.731–16.805 ms to 0.301–0.313 ms including fresh lexing. Depth-100 batches improve in both query modes. Small fresh-token controls are mixed: a depth-five single terminal query spans 0.011 ms baseline and 0.011–0.016 ms fixed. Plain assignment controls are similar or mixed. No universal whole-analyzer speedup is claimed.

These are syntactic editing/stress fixtures. They do not establish that Office compiles arbitrarily deep or long physical lines; no Office execution or VS Code UI latency was measured.

Reproduce from the repository root:

```powershell
node scripts/benchmark-reference-target-parens.mjs --shape=inline-if --baseline=0f685d45
node scripts/benchmark-reference-target-parens.mjs --shape=inline-if
node scripts/benchmark-reference-target-parens.mjs --shape=inline-if
node scripts/benchmark-reference-target-parens.mjs --shape=inline-if --baseline=0f685d45
npm exec vitest run tests/referenceInlineIfWork.test.ts tests/vbaReferenceKinds.test.ts tests/referenceTargetParenWork.test.ts
npm run check-types
npm test -- --run
```

## Validation

Twelve new tests (four work tests fail on baseline), 39 focused tests, type checking and the full suite pass. Full suite: 654 files, 13,254 tests passed and 13 skipped on main0f685d45 plus this runtime change. Later edits only extend the standalone benchmark and add documentation.

Complete ordered-map comparison against baseline covers all 8,256 corpus sources in four offset modes (33,024 comparisons), 5,832 generated parenthesis/assignment cases, and 37,620 generated inline cases: 76,476 comparisons total, all matching. Inline combinations include three conditional heads, eleven valid/malformed conditions, nineteen leaf statement families/tails, five Else suffixes and three newline forms, with sparse, all, reversed/duplicate and unusual/default offsets. Both bundles freeze actual cached full-source token objects/arrays; no uncaught exceptions occurred. This is token ownership evidence, not frozen AST or complete repository coverage.
