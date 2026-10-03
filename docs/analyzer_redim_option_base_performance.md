# ReDim Preserve module option setup

The ReDim Preserve rule previously resolved Option Base separately for each active procedure. When no option was present, each lookup visited every active module member. This added quadratic work in procedure count before checking the actual ReDim statements.

Resolve the option lazily once per rule invocation, after finding the first active procedure. Reuse the value for its remaining procedures and nested bodies. Keep the value local so conditional-compilation environments cannot share stale facts. A module without active procedures does not perform option setup.

## Reproduction

```powershell
node scripts/benchmark-redim-option-base.mjs --baseline=4efed8bb --rounds=15
node scripts/benchmark-redim-option-base.mjs --rounds=15
```

The benchmark times the public ReDim Preserve rule after parsing, with warmed token caches, three warmups and 15 samples. Each short procedure declares a dynamic array, sets bounds 1 To 3, and preserves it with an implicit lower bound. The default-base fixture produces one diagnostic per procedure; the Option Base 1 control produces none. The harness asserts these counts.

Node 24.18.0, AMD Ryzen 7 9800X3D; median milliseconds:

| Procedures | Default base before | After | Explicit Base 1 before | After |
| ---: | ---: | ---: | ---: | ---: |
| 100 | 0.216 | 0.212 | 0.195 | 0.213 |
| 1,000 | 2.630 | 1.253 | 1.245 | 1.176 |
| 3,000 | 17.482 | 3.637 | 4.253 | 3.441 |

These isolated-rule stress measurements exclude parsing and do not represent editor latency. Absolute timings vary with load. Explicit Option Base at the module start was already a cheap lookup and remains a control.

## Validation

- 86 targeted tests passed, covering the new setup budget and existing array diagnostics.
- 2,000 generated fixtures had identical complete diagnostic callback outputs against baseline 4efed8bb. Cases cover default/explicit/conditional options, dimension counts and bounds, unknown expressions, erasures, nested blocks, and multiple procedures.
- Member-kind reads for 200 procedures without Option Base fell from 40,200 to 400. The regression test requires linear work rather than wall-clock timing.
- Regression cases verify Base 0/1 across all procedures, activity changes on the same parsed module, and no option-text reads when no procedures are active.

- Type checking passed; full suite: 592 files, 12,443 tests passed, 13 skipped.
