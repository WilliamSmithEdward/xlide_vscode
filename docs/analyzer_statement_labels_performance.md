# Deferred statement-form label collection

Statement-form diagnostics collected label declarations and references and built an offset Set for every executable span. Labels only affect this rule when a bare identifier survives qualification/shadow checks and matches another standard module of the project. Collection is now deferred to that point and shared for all matching tokens in that span. The missing-token guard and label namespace remain; no persistent cache or dependency is added.

An actual-rule probe counts 1,000 declaration-helper calls and 1,000 reference-helper calls for 1,000 ordinary assignments, reduced to zero. The same zero-work behavior is tested with unrelated project metadata, qualified names and locals shadowing a module name. A two-hit span still shares one collection, while a label declaration and GoTo target retain their exclusions.

## Reproduction

```powershell
node scripts/benchmark-analyzer-statement-labels.mjs --baseline=ae6715562299e15f17452cf7332e00ed677ef429
node scripts/benchmark-analyzer-statement-labels.mjs
node scripts/benchmark-analyzer-statement-labels.mjs --rule-only --baseline=ae6715562299e15f17452cf7332e00ed677ef429
node scripts/benchmark-analyzer-statement-labels.mjs --rule-only
```

Each mode ran sequential before/after, without other audit tests/benchmarks running. esbuild bundles the actual code; only statementForms.ts is substituted for the baseline. Fifteen measured rounds after three warm-ups; v24.18.0, AMD Ryzen 7 9800X3D 8-Core Processor. Times are milliseconds per pass. No work counters run inside timed samples. Complete-analyzer mode parses/tokenizes outside timing, uses warm sources or unique trailing comments for fresh parsed bodies (body identity checked), and includes the full analyzeModule pass. Host globals remain warm. Rule-only mode parses and builds symbols outside timing, measuring the actual checkStatementForms pass. Fresh rule samples likewise change source text with trailing comments. Expected diagnostic code/message/spans are checked outside timing in both modes; complete-analyzer internal errors fail the sample.

## Complete analyzer

| Fixture | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: |
| many-warm | 48.079 | 54.187 | 60.388 | 87.829 |
| many-fresh | 53.822 | 54.487 | 67.128 | 77.286 |
| body-warm | 111.722 | 115.886 | 191.125 | 130.417 |
| body-fresh | 131.231 | 137.117 | 141.985 | 216.415 |
| branches-warm | 57.425 | 48.807 | 82.495 | 59.006 |
| branches-fresh | 58.857 | 49.047 | 96.321 | 54.288 |
| types-warm | 23.636 | 23.061 | 26.394 | 51.972 |
| types-fresh | 24.893 | 24.858 | 29.807 | 29.341 |
| loops-warm | 35.359 | 40.377 | 42.600 | 46.751 |
| loops-fresh | 40.814 | 38.692 | 58.253 | 45.446 |
| short-warm | 0.440 | 0.347 | 0.452 | 0.489 |
| short-fresh | 0.324 | 0.324 | 0.450 | 0.471 |
| qualified-warm | 9.334 | 9.070 | 50.642 | 10.685 |
| qualified-fresh | 10.021 | 13.034 | 11.189 | 45.001 |
| label-hit-warm | 19.885 | 21.095 | 24.424 | 26.281 |
| label-hit-fresh | 21.044 | 23.350 | 24.951 | 27.439 |

Fixtures cover many procedures, one long body, branches, declaration-only types, loops, a short module, qualified project-module names and real label/module-name hits.

## Statement rule alone

| Fixture | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: |
| assignments-warm | 0.616 | 0.697 | 1.288 | 3.392 |
| assignments-fresh | 0.503 | 0.452 | 2.180 | 1.244 |
| long-body-warm | 1.705 | 0.948 | 4.132 | 1.518 |
| long-body-fresh | 4.383 | 1.943 | 5.909 | 3.423 |
| project-no-hit-warm | 0.440 | 0.212 | 0.626 | 0.289 |
| project-no-hit-fresh | 0.781 | 0.335 | 0.956 | 0.413 |
| qualified-warm | 16.948 | 18.548 | 25.392 | 22.318 |
| qualified-fresh | 16.539 | 18.028 | 21.114 | 22.288 |
| label-hit-warm | 0.659 | 0.691 | 0.932 | 1.157 |
| label-hit-fresh | 0.746 | 0.741 | 1.318 | 2.634 |
| short-warm | 0.003 | 0.003 | 0.009 | 0.008 |
| short-fresh | 0.004 | 0.004 | 0.010 | 0.013 |

Rule fixtures use 1,000 assignments, 5,000 assignments, 1,000 assignments with unrelated project metadata, 1,000 qualified names, 200 repeated label/jump/module-value groups, and one assignment.

This is a confirmed unnecessary-work reduction, not a universal analyzer latency improvement. Long-body rule medians improve (warm 1.705 -> 0.948 ms, fresh 4.383 -> 1.943), as do unrelated-project rule cases. Whole-analyzer timings are mixed: branch warm 57.425 -> 48.807, but many-procedure warm 48.079 -> 54.187, body warm 111.722 -> 115.886 and label-hit fresh 21.044 -> 23.350. Rule-only qualified warm also rises 16.948 -> 18.548. Other rules dominate these complete-pass fixtures, and these measurements do not establish a net speedup for every workload. All controls/p95 values are retained above.

No actual editor interaction latency was measured.

A native differential run compared complete diagnostic objects and internal-error lists for all 8,256 string-source oracle cases across four hosts, both default and other-standard-module contexts: 66,048 comparisons. Cached token arrays/objects were frozen in both bundles. Every output matched baseline ae671556, with no internal errors. Non-string corpus entries are outside this differential scope.
