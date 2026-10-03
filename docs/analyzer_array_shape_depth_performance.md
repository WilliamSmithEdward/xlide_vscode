# Array-shape nesting safety

arrayValueShape recursively follows nested Array elements and Filter source arguments. It previously had no nesting budget. A direct 2,000-level Array probe threw a JavaScript stack overflow; a 5,000-level probe also threw. Deep generated or incomplete editor expressions must not crash this path.

The shape builder now shares MAX_EXPRESSION_DEPTH (256) with expression parsing and numeric folding. A private recursive helper carries one budget through both Array and Filter calls; the public signature is unchanged. At the boundary, deeper facts become unknown. Outer Array dimensions remain known; a Filter whose input cannot be evaluated returns unknown. This intentionally stops producing complete shape facts beyond the limit, including depths the previous implementation could sometimes finish before overflowing.

## Reproduction

```powershell
node scripts/benchmark-array-shape-depth.mjs --baseline=0233b9bd --rounds=15
node scripts/benchmark-array-shape-depth.mjs --rounds=15
```

The benchmark times the public shape builder with tokens prepared outside the timed region. It uses three warmups and 15 samples for nested Array calls and nested Filter calls with a known binary comparison. It reports stack-error counts and returned shape depth. These deliberately oversized expressions are analyzer stress inputs, not a claim about valid VBA nesting limits or editor latency. Timing thrown calls records time until failure, not successful equivalent work.

Node 24.18.0, AMD Ryzen 7 9800X3D; median milliseconds:

| Nesting | Array before | After | Filter before | After |
| ---: | ---: | ---: | ---: | ---: |
| 10 | 0.018 | 0.014 | 0.024 | 0.024 |
| 100 | 0.328 | 0.312 | 0.471 | 0.478 |
| 1,000 | 35.697 | 14.297 | 41.466 | 16.096 |
| 5,000 | 562.291 (throws) | 86.718 | 1,407.533 (throws) | 181.214 |

Both 5,000-level baseline fixtures threw on all 18 calls, including warmups. The fixed version had no stack errors. At 1,000 and 5,000 levels, Array returns 256 outer shapes; Filter returns unknown. Ordinary 10- and 100-level results stay unchanged. The guard bounds recursion, while token processing still scales with input size and the allowed nesting depth. Absolute timings vary with load; an initial single-call 5,000-level Array probe threw after 397.256 ms.

## Validation

- Type checking passed; full suite: 595 files, 12,516 tests passed, 13 skipped.
- 86 targeted tests passed.
- 3,000 generated shapes at nesting up to 80 matched baseline outputs exactly, covering Array/VBA.Array, Filter, Split, Range values, literals, unknown and malformed inputs, Option Base and comparison modes.
- Five regressions cover nested bounds and values below the budget, its exact boundary, 5,000 nested arrays, 2,000 nested Filters and mixed recursion.
