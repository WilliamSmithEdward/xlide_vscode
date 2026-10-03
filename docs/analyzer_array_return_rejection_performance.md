# Rejected array-return summaries

arrayOnlyVariantFunctions classifies Variant-returning Functions whose every use of their return variable is a whole Array(...) assignment. A scalar assignment, self-read or incompatible statement permanently disqualifies the function. The old callback kept tokenizing/checking later spans even after onlyArrays was false.

Return immediately within the rejecting statement and skip inspection callbacks after rejection. The existing statement walker and conditional-activity checks still visit the body; this is not a constant-time traversal claim. Valid array-only, late rejection and never-assigned Empty cases keep their existing classification. No cache, public API or dependency is added.

## Reproduce

```powershell
node scripts/benchmark-array-return-rejection.mjs --baseline=205e335c9132de3071a6c7a9d722a6c81a10ea4b
node scripts/benchmark-array-return-rejection.mjs
```

The harness exposes the private summary only in its scratch bundle, never in the production API. It substitutes only assignments.ts for the baseline. Parse once, build fresh bound symbols outside each timer (assert distinct root identity), and separately time the summary and the full checkAssignmentTypes rule. Three warmups, 15 samples, median/p95; run without other tests or benchmarks in parallel.

Each fixture has a Variant Function and the stated number of Debug.Print n statements, plus an early scalar assignment/self-read, a late scalar assignment, or only an Array assignment. A caller assigns F() to a dynamic Long array. The summary must include F only for the array-only mode; that mode must report one assignment mismatch and the other modes none. These are synthetic rule/helper timings, not overall analyzer latency.

## Measurements

2026-10-03, Node v24.18.0, AMD Ryzen 7 9800X3D 8-Core Processor. Medians in milliseconds:

| Body statements, mode and timed surface | Before | After |
| --- | ---: | ---: |
| 100-early-scalar-summary | 0.020 | 0.008 |
| 100-early-scalar-rule | 0.075 | 0.089 |
| 100-early-self-read-summary | 0.014 | 0.002 |
| 100-early-self-read-rule | 0.192 | 0.190 |
| 100-late-scalar-summary | 0.023 | 0.028 |
| 100-late-scalar-rule | 0.135 | 0.141 |
| 100-only-arrays-summary | 0.008 | 0.010 |
| 100-only-arrays-rule | 0.082 | 0.072 |
| 1000-early-scalar-summary | 0.079 | 0.008 |
| 1000-early-scalar-rule | 0.324 | 0.261 |
| 1000-early-self-read-summary | 0.055 | 0.005 |
| 1000-early-self-read-rule | 0.581 | 0.638 |
| 1000-late-scalar-summary | 0.057 | 0.083 |
| 1000-late-scalar-rule | 0.446 | 0.468 |
| 1000-only-arrays-summary | 0.051 | 0.053 |
| 1000-only-arrays-rule | 0.227 | 0.222 |
| 10000-early-scalar-summary | 0.555 | 0.039 |
| 10000-early-scalar-rule | 2.254 | 1.803 |
| 10000-early-self-read-summary | 0.761 | 0.039 |
| 10000-early-self-read-rule | 4.159 | 3.788 |
| 10000-late-scalar-summary | 0.555 | 0.561 |
| 10000-late-scalar-rule | 4.301 | 4.339 |
| 10000-only-arrays-summary | 0.560 | 0.589 |
| 10000-only-arrays-rule | 2.068 | 2.193 |

For 10,000 following statements, early scalar summary time is 0.555 to 0.039 ms and early self-read 0.761 to 0.039 ms. The corresponding rule times are 2.254 to 1.803 and 4.159 to 3.788 ms. Controls and smaller whole-rule cases are mixed: late rejection 4.301 to 4.339 ms; valid array-only 2.068 to 2.193 ms. Earlier probes varied by roughly a factor of two and sometimes showed whole-rule regressions on early rejection. The reliable claim is fewer unnecessary summary inspections; no general analyzer speedup is claimed.

An instrumented scratch bundle counts inspected statement spans independently of timing: early scalar 10,001 to 1; late scalar 10,002 unchanged; array-only 10,001 unchanged. Variable declaration groups are not spans this statement walker inspects. All original traversal/activity checks remain.

## Validation

- Four tests preserve irreversible rejection, self-read/single-line conditions, array-only and never-assigned Empty cases, and conditional activity on reused parsed nodes.
- Three focused files: 198 tests passed; typecheck passed.
- Full suite: 608 files, 12,640 tests passed, 13 skipped.
- 1,500 complete analyzeModule outputs exactly match baseline 205e335c with no internal errors. Generated cases combine scalar/Array/Set/self/field/ByRef/loop/conditional/single-line uses, Variant/Long/suffixed returns, four caller array element types and conditional activity. Messages, spans and order match.
