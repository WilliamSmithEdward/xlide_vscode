# Initializer name membership

Introduce Parameter rejects an initializer that refers to the procedure’s locals/parameters or module-private names. Its first-occurrence output was deduplicated by searching the growing output array for each name. With distinct blocked names, this revisits every earlier entry.

A lazy, query-owned Set tracks exact spellings. It preserves first-occurrence ordering, case-sensitive output deduplication and case-insensitive binding checks. Queries without a blocked name do not allocate the set; function-result handling remains on its existing path. No persistent cache or external-state memoization is added.

## Validation

Three work tests fail on baseline; five semantic controls pass. After repair, type checking, 107 focused tests and the full suite pass: 777 files / 15,201 tests, with 7 files / 33 tests skipped, in 74.46 seconds. Exact refusal controls include repeated names, distinct capitalization, strings and unbound identifiers. The shared continuation-reader fix in PR1154 is a dependency so large, physically valid inputs reach this path.

Against `6c172ee3206cc2028e879941dd0c19c7877dc7f1`, 20,000 seeded complete refactor queries with mixed public/private declarations and repeated spellings preserve full results and edited source. Two public refactors also preserve 49,536 complete result/edited-source/cross-module-edit queries on single-line target initializers in all 8,256 oracle source environments across three line endings. No differences were observed; the oracle comparison is environment compatibility evidence, not a broad proof of semantic validity.

## Work and timings

At 1,000 distinct blocked names, the growing-array lookup searches 499,500 prior slots; at 2,000 it searches 1,999,000. The repaired path performs no such searches. The set still does one membership lookup and insertion per new blocked name. Counts are work evidence, not heap bytes.

Run `node scripts/benchmark-initializer-name-membership.mjs`, optionally with `--baseline=6c172ee3206cc2028e879941dd0c19c7877dc7f1`. Node v24.18.0 on AMD Ryzen 7 9800X3D 8-Core Processor           . Sequential ABBA runs use three warmups and nine samples. The clock covers the complete public refactor call with a warmed module AST; every complete refusal reason is independently specified. The 2,000-name fixture uses 20 physical expression lines / 19 continuations, with maximum physical line length below 1,023. Work instrumentation runs before timings and affects JIT state.

| Blocked names | Before median ms | After median ms |
|---:|---:|---:|
| 1 | 0.0207–0.0233 | 0.0231–0.0235 |
| 100 | 0.0548–0.0569 | 0.0457–0.0478 |
| 1000 | 1.1183–1.1192 | 1.0803–1.0816 |
| 2000 | 2.6539–2.882 | 0.9837–1.0751 |

The 2,000-name fixture improves; one-name measurements overlap, and the 1,000-name gain is small. Earlier narrow harness runs had overlapping 1,000-name results; these measurements do not establish a general refactor/editor speedup. This benchmark measures a refusal path, not successful parameter migration. No cold-parser or heap improvement is claimed.
