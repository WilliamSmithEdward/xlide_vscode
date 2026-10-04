# Significant-token consumers across the analyzer

The shared statement-token helpers return significant tokens, excluding comments/newlines in both the cached derivation and fallback lexer. Leading-label and line-number wrappers preserve that property. Analyzer consumers nevertheless filtered those arrays again, copying tokens and checking kinds that cannot occur in the input.

An import- and predicate-checked AST audit identifies 95 sites across 46 analyzer files. Remove 92 filters containing only comment/newline exclusions, and simplify three mixed predicates while retaining integer-literal or Else exclusions. The changes cover shared callee/dataflow/function/held/loop facts and diagnostic families. No new helper or cache is added. Inputs continue to be read without mutation, required slices remain, and filters on arbitrary token inputs are unchanged.

The initial line-oriented inventory found 94 candidates; AST inspection includes multiline sites and distinguishes functional mixed predicates. The cached-token contract is verified at the lexer and imported wrappers, candidate bindings have no direct array mutators, and full-analyzer ownership checks freeze the actual cached arrays and token objects to catch downstream helper/alias writes. This validation establishes the edited paths, not exhaustive branch coverage of the repository.

## Deterministic cost

Scratch-only instrumentation counts actual redundant filter calls/copies and comment/newline predicate checks. These fixtures use Excel; equivalent output/ownership checks also run under Word, PowerPoint and Access:

| Fixture | Redundant calls before | Fully redundant array copies before | Kind checks before | Redundant calls/copies/checks after | Diagnostics |
| --- | ---: | ---: | ---: | ---: | ---: |
| 100 procedures, 12 writes each | 29,001 | 29,001 | 144,704 | 0 | 100 |
| One procedure, 1,000 writes | 24,003 | 24,003 | 120,011 | 0 | 1 |
| 20 branch/array/Collection procedures | 4,941 | 4,941 | 22,024 | 0 | 0 |
| 100 procedures with typed declarations | 101 | 101 | 504 | 0 | 100 |
| Labels/files/dictionary/error/Mid edge program | 509 | 507 | 2,792 | 0 | 2 |
| 10 loop/Select procedures | 2,121 | 2,121 | 7,994 | 0 | 0 |

Two edge-program calls use mixed predicates, so their copies are functional and retained; their redundant kind checks disappear. Zero refers to redundant filters/checks, not all analyzer allocations. There are no production or timing counters. Frozen-cache probes for these six fixtures across four hosts preserve complete outputs with no internal errors.

## Reproduction and timings

From the repository root with dependencies installed:

```powershell
node scripts/benchmark-analyzer-significant-token-consumers.mjs --baseline=f999f6b1
node scripts/benchmark-analyzer-significant-token-consumers.mjs
```

The benchmark replaces only the 46 changed production files from f999f6b1 in the current harness. The baseline includes the prior assignment-token cleanup. Timing has no counters or frozen-cache instrumentation.

Fixtures: many has 150 procedures with 12 arithmetic writes each; body has one procedure with 5,000 writes; branches has 100 procedures with If/ReDim/array/Collection/For behavior; types has 500 procedures with typed Range declarations; loops has 100 procedures with For/Do/Select; short has one simple assignment. All use Excel. Every sample checks diagnostic codes/messages/spans and rejects internal errors outside the timed interval.

Warm mode repeats source with reused parsed bodies. Fresh mode appends a unique trailing comment per sample and asserts a fresh parsed procedure body. Parsing and whole-source tokenization happen before timing; the complete analyzeModule call, including its binding and rule work, is timed. Fresh body-dependent analysis is not prewarmed, while host/global metadata and process startup are already warm. These are synthetic complete-analyzer timings, not editor latency or cold process startup. Three warmups precede 15 samples. Before/after runs were sequential after tests and corpus/counter probes finished, without concurrent audit workloads.

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor           ; measured locally on 2026-10-03. Times in milliseconds.

| Fixture / mode | Diagnostics | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: | ---: |
| many-warm | 150 | 49.327 | 48.464 | 62.198 | 61.602 |
| many-fresh | 150 | 57.653 | 58.980 | 79.380 | 74.428 |
| body-warm | 1 | 120.866 | 116.468 | 126.138 | 126.741 |
| body-fresh | 1 | 139.528 | 136.701 | 176.917 | 185.417 |
| branches-warm | 0 | 41.264 | 39.707 | 49.618 | 43.287 |
| branches-fresh | 0 | 45.657 | 41.401 | 72.636 | 45.230 |
| types-warm | 500 | 20.910 | 19.806 | 22.174 | 20.659 |
| types-fresh | 500 | 20.925 | 21.102 | 39.296 | 22.778 |
| loops-warm | 0 | 31.039 | 31.392 | 36.030 | 36.212 |
| loops-fresh | 0 | 35.102 | 35.224 | 36.511 | 40.593 |
| short-warm | 1 | 0.300 | 0.307 | 0.381 | 0.353 |
| short-fresh | 1 | 0.312 | 0.326 | 0.335 | 0.435 |

The table reports the complete integrated-tree measurements, including controls and p95 values. Results vary by fixture and cache mode; the deterministic redundant copies are removed, but other allocation, binding, traversal and dataflow work remains. These measurements do not establish a universal latency improvement.

## Validation and remaining finding

- Type check passed.
- 45 focused ownership/file/late-bound/assignment tests passed before integration; 36 ownership/file/late-bound/doc-closing tests passed on the integrated tree, including nine new full-analyzer frozen-cache regressions. The functional Else filter is exercised with the actual project file-state option.
- Full suite: 628 files, 12,869 tests passed, 13 skipped.
- 24 instrumented frozen-cache fixture/host comparisons preserve complete outputs with no internal errors.
- Native deep comparison of complete diagnostics on 8,229 source-bearing oracle cases across four hosts passes: 32,916 analyzer runs, with every cached statement array/token frozen in both bundles. All diagnostics and internal-error lists match. The remaining 816 oracle records lack a single source field and are not covered by this comparison.
- Eight runs on two reserved-label corpus cases retain the identical pre-existing statementForms exception in both bundles (issue272_01_compile and issue272_02_compile under four hosts). No new or changed internal errors occur. That separate missing-token failure is filed as issue #853, which also covers bare Call; it is not silenced or attributed to this cleanup.

Generated AST plans, ownership/counter probes and corpus comparisons remain audit scratch material. The complete-analyzer benchmark and focused ownership regressions are committed. The broader analyzer and repository audit remains ongoing.
