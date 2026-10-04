# Extract Method local sort positions

Extract Method sorted touched locals by searching the module source for both names on every comparison. The private LocalUse record now captures that same raw first-occurrence position once. Zero or one touched local needs no ordering search. Parameter order, moved declaration order, comment/case behavior and stable substring ties keep the original comparator's semantics; no AST or public result field changes.

A regression through actual extractMethod counts 15,062 matching source searches for 1,000 touched locals before, and 1,000 after. Three baseline compatibility controls cover comment order, stable ties, procedure parameters and a single local.

## Reproduce

Baseline 804727259fd619e51f3b8bd58f5be4816405c51f; Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor           . Reuse the existing benchmark, extended with baseline substitution and optional local-order workloads. Only src/analyzer/refactor/extractMethod.ts is replaced in the baseline bundle; other code and workload generation are identical.

```powershell
node scripts/benchmark-extract-method.mjs --baseline=804727259fd619e51f3b8bd58f5be4816405c51f --local-order > before.json
node scripts/benchmark-extract-method.mjs --local-order > after.json
```

Each workload has 15 measured rounds after three warm-ups. These are complete extractMethod timings, including its parse/reference work. Warm mode reuses the same source; fresh mode adds a unique trailing comment, so the full-source cache key changes. Input construction, result assertions and hashing are outside the clock. Padding and raw name-order comments are split across bounded physical lines. Before/after runs are sequential, with no own concurrent tests or probes. Source caches may retain reusable internals; fresh mode does not promise that every internal cache is cold.

The script checks the hash of the complete refactor result on every run. All 16 before/after hashes matched. The read-heavy large cases are synthetic stress workloads and their generated parameter lists are not checked by the VBE; the 3,000-local write case moves declarations instead of creating parameters. No Office or VS Code UI latency was measured. Small/mostly-untouched controls are mixed; no universal speedup is claimed.

## Results

Milliseconds per extraction. Read workloads touch every local; write workloads write each local first and move its declaration. The few workload touches only five of 1,000 declarations. Padding is the number of extra comment characters before the raw first occurrences.

| Locals | Touched | Padding | Order | Form | Source | Median before | Median after | p95 before | p95 after |
| ---: | ---: | ---: | --- | --- | --- | ---: | ---: | ---: | ---: |
| 1 | 1 | 0 | ordered | read | warm | 0.020 | 0.020 | 0.043 | 0.044 |
| 1 | 1 | 0 | ordered | read | fresh | 0.079 | 0.074 | 0.374 | 0.149 |
| 5 | 5 | 0 | shuffled | read | warm | 0.020 | 0.018 | 0.039 | 0.036 |
| 5 | 5 | 0 | shuffled | read | fresh | 0.066 | 0.069 | 0.320 | 0.312 |
| 100 | 100 | 0 | shuffled | read | warm | 0.457 | 0.251 | 0.609 | 0.449 |
| 100 | 100 | 0 | shuffled | read | fresh | 1.040 | 0.819 | 1.567 | 1.187 |
| 1000 | 1000 | 0 | ordered | read | warm | 5.833 | 3.870 | 6.880 | 5.298 |
| 1000 | 1000 | 0 | ordered | read | fresh | 9.815 | 8.214 | 12.663 | 9.956 |
| 1000 | 1000 | 8000 | shuffled | read | warm | 35.686 | 3.991 | 41.292 | 5.469 |
| 1000 | 1000 | 8000 | shuffled | read | fresh | 40.129 | 6.392 | 55.738 | 9.629 |
| 1000 | 1000 | 100000 | shuffled | read | warm | 56.913 | 4.918 | 65.344 | 6.273 |
| 1000 | 1000 | 100000 | shuffled | read | fresh | 69.165 | 8.462 | 77.014 | 10.567 |
| 3000 | 3000 | 100000 | shuffled | write | warm | 429.539 | 27.668 | 449.819 | 33.000 |
| 3000 | 3000 | 100000 | shuffled | write | fresh | 439.752 | 38.414 | 458.958 | 46.342 |
| 1000 | 5 | 100000 | shuffled | few | warm | 0.559 | 0.511 | 0.999 | 1.725 |
| 1000 | 5 | 100000 | shuffled | few | fresh | 4.247 | 3.020 | 8.461 | 4.677 |

Validation: 20 focused tests and type checks pass. Native baseline/fixed comparison of 600 generated cases has identical complete results and applied text (327 successful extractions, 273 matching refusals), with both parsed ASTs deeply frozen. Cases cover LF/CRLF/CR, Unicode and mixed case, raw comments, parameters, read/write/ByRef/output forms, Static locals, partial selections and refusals. This is pure refactor output/ownership parity, not VBE execution. Default benchmark mode smoke-tested with three rounds. Full suite passes: 639 files, 12,986 tests passed, 13 skipped.
