# Prepared cross-module constant dependency walks

Cross-module constant-cycle checks previously lexed and resolved a dependency expression each time a starting constant reached it. The recursive walk also consumed a JavaScript stack frame for every dependency, throwing RangeError on a 10,000-node cycle.

Prepare dependency edges once per expression/module scope within the rule pass. Replay each starting constant with explicit DFS frames, preserving source reference order, the seen-node rules and the first foreign reference named in a cycle diagnostic. The per-start graph traversal remains; this change removes repeated expression preparation and call-stack depth. There is no persistent cache.

## Reproduction

Run node scripts/benchmark-declaration-dependencies.mjs --rounds=15 against the target checkout and baseline 76bf7053. The runner bundles local source with the existing esbuild dependency, parses fixtures before timing and reports median/p95 over 15 samples after three warmups. Each module root references the same external chain ending in a numeric value; no diagnostics are expected. No test suite ran concurrently with either measurement.

Node 24.18.0, AMD Ryzen 7 9800X3D, Windows; milliseconds:

| Fixture | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: |
| 1-roots/100-external-chain | 0.132 | 0.115 | 0.337 | 0.254 |
| 100-roots/100-external-chain | 3.427 | 1.256 | 5.428 | 1.492 |
| 200-roots/500-external-chain | 33.452 | 12.829 | 39.932 | 14.227 |

The large shared-chain fixture is approximately 62% faster. These are warm synthetic rule-pass timings excluding parsing, not whole-workbook latency guarantees.

## Validation

The shared-expression regression reproduces 10,000 tokenize calls on original code (budget below 300). A 10,000-node cyclic graph reproduces Maximum call stack size exceeded on original code and produces the expected cycle diagnostic after the fix. Another regression verifies first foreign-reference provenance, then changes the project map and requires updated results in a new pass. All 25 existing declaration-order tests pass.
Exact diagnostic code/message/span outputs also matched on 300 seeded cyclic graphs. Type checking and the full Vitest suite passed: 562 files, 12,003 tests passed and 13 skipped.
