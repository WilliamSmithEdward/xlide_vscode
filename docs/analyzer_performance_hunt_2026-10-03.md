# Analyzer performance hunt - 2026-10-03

Base: `3dec1b4cea2f86e0dfe7cb8c8fe02a26ae35c22e` (main).
Environment: v24.18.0, Windows, AMD Ryzen 7 9800X3D 8-Core Processor.

## Method

Run `node scripts/benchmark-analyzer.mjs --rounds=15`.
The script bundles the checked-out analyzer with the existing esbuild dependency,
warms each fixture three times, and reports the median and empirical p95 of
15 samples. Before and after use the same script and synthetic sources. Final
measurements were collected after the test workload finished. These are local
microbenchmarks, not VS Code latency measurements or cross-machine budgets.

Full analysis uses `analyzeVbaModuleSource`. Body edits use the incremental
semantic engine with a prior state and one changed assignment. Full synthetic
analysis must produce no diagnostics or internal errors. The literal setup
case builds fresh symbols for each sample and requests the same procedure's
value view twenty times, representing independent rules. Class facts are
checked to contain every generated getter. The warm project-index case is a
control: these changes do not optimize its visibility queries.

## Measurements

All times below are medians in milliseconds. Positive reduction means faster.

| Fixture / path | Before | After | Reduction |
|---|---:|---:|---:|
| many-100/full-warm | 55.755 | 47.432 | 14.9% |
| many-100/body-edit | 33.475 | 27.164 | 18.9% |
| many-300/full-warm | 171.22 | 136.892 | 20.0% |
| many-300/body-edit | 112.649 | 92.604 | 17.8% |
| body-6000/full-warm | 214.737 | 156.944 | 26.9% |
| body-6000/body-edit | 192.629 | 156.58 | 18.7% |
| numeric-type/1000000 | 125.303 | 3.101 | 97.5% |
| literal-setup/20-rules | 4.369 | 0.476 | 89.1% |
| malformed-headers/100 | 1.313 | 1.297 | 1.2% |
| class-member-facts/100 | 0.818 | 0.157 | 80.8% |
| malformed-headers/500 | 19.541 | 4.405 | 77.5% |
| class-member-facts/500 | 22.76 | 0.709 | 96.9% |
| malformed-headers/1000 | 89.923 | 10.27 | 88.6% |
| class-member-facts/1000 | 66.639 | 0.652 | 99.0% |
| project-index/50-modules-warm | 1.437 | 1.566 | -9.0% |

The 300-procedure fixture has 5,701 lines. The large single-procedure fixture
has 6,008 lines. The parser's warm cache, lexer/span-token reuse, a real VB6
module, and branch-heavy code were also sampled during the CPU-profile pass;
no new regression was established for those paths. Earlier project visibility
and ElseIf/local-count performance reports (#173 and #322) are already closed.

## Findings and fixes

1. `isNumericType` rebuilt a nine-entry Set for every classification. Reuse
   one private constant Set, preserving the existing normalized-name contract.
2. `knownLocalLiteralValuesAt` rebuilt the procedure-wide literal scan and
   per-statement derived views for independent rules. Share one view per
   symbols/procedure identity, guarded by source and conditional-activity
   identity. Weak keys bound its lifetime to that analysis context.
3. `checkMalformedLines` repeatedly searched the module's whole token stream
   for each header and the whole procedure list for each statement's location.
   Binary-search the ordered token and non-overlapping member spans instead,
   preserving the strict span-boundary comparisons and unnamed-procedure rules.
4. `classMemberValues` scanned all tokens for every field/getter. Index each
   identifier's first/last occurrence once for field checks and binary-search
   each getter's body range. Preserve the previous conservative mention checks
   and token filtering, including comments and return-value forms.

## Validation

- `npm run check-types` and `git diff --check`.
- Full suite: 549 files passed; 11,950 tests passed, 13 skipped.
- Focused diagnostics/incremental/module-state suite: 125 tests passed.
- New regressions cover exact token boundaries, logarithmic token lookup,
  repeated rule view reuse and invalidation, per-statement values, malformed
  trailing headers, class-fact semantics, and token-read complexity.

Timing assertions are intentionally absent from CI; token-read counts guard
algorithmic complexity without depending on machine speed.
