# Qualified host type performance audit — 2026-10-03

## Finding

Qualified type resolution rebuilt the library list and host candidates for every declaration. The bare-name path was already indexed. Qualified completion also enumerated the entire model per request.

The fix groups immutable host metadata once per model identity, then uses a case-insensitive library/type lookup. Documentation is rendered only for queried types. Qualified completion reuses the grouping, with new result records so callers cannot mutate the index. Project-module types retain priority, including ambiguous results, followed by stdole and host types. First library spelling, type spelling, duplicates and insertion order are preserved; nested type keys and enums lacking library keys remain excluded.

## Measurements

Run `node scripts/benchmark-qualified-types.mjs` in each checkout. Node 24.18.0, AMD Ryzen 7 9800X3D, three warmups and 15 samples; median milliseconds, no concurrent tests. Baseline: 3dec1b4. The comparison bundles the baseline typeCompletion.ts with identical unchanged analyzer dependencies.

| Workload | Before | After |
| --- | ---: | ---: |
| 1,000 qualified Excel.Range lookups | 62.291 | 0.351 |
| 100 qualified completion requests | 7.249 | 1.961 |
| Color 100 qualified declarations | 6.084 | 0.054 |
| Color 500 qualified declarations | 30.352 | 0.155 |
| Color 2,000 qualified declarations | 124.460 | 0.925 |

These are warm synthetic measurements. Building a new model index still costs one metadata scan; results do not predict end-to-end editor latency for every project.

## Validation

Focused tests cover enumeration counts across hits/misses/completions, case variants, duplicates, order, nested keys, project and stdole precedence, ambiguous project types, model replacement and caller mutations. Full suite: 549 files, 11,933 passed and 13 skipped. Type checking and extension compilation passed.

## Continuing audit

The analyzer inventory contains 183 TypeScript files across 21 directories. The first performance PR fixed numeric classification, repeated local-literal setup, malformed-line lookup and class-member scans. This pass measured lexer/parser, symbol construction, type references, hover, signature help, four completion paths, three semantic paths, reference classification, formatting, expression resolution, refactoring, call sites, documentation, conditional compilation, annotations and constants.

This is ongoing coverage, not a claim of an exhaustive completed audit. Next checks include per-rule diagnostics, symbol association for DefType and member attributes, large-module formatting, and remaining refactoring paths. Exact duplicated token walkers are being reviewed separately; supported embedding APIs and intentional test seams are retained despite low production reference counts. The repository-wide audit follows the analyzer pass.
