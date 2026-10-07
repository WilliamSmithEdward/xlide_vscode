# Inference cache source adoption

Three inference caches (`knownLocalLiteralValuesAt`, `unreachableStatementsIn`,
and `functionResultFor`) retain results within bound-symbol and procedure
identities. Source text and conditional activity must match; function calls also
retain separate entries for their argument tokens and object-result mode.
Previously, equal-text hits kept the old source string and could compare the
complete module on every following lookup. Matching result hits now adopt the
caller string. Results, identity, context keys and invalidation are unchanged.

## Reproduction

Run `node scripts/benchmark-inference-source.mjs`, or add
`--baseline=7d8bba83da6513da66b56a5ad1f6e2029e06ecfe` for the prior implementation.
The script checks independently expected local values, dead statements and
function result tokens before timing, plus cached identity on every hit.
Each of three consumers is measured at approximately 1 KB, 100 KB and 1 MB,
with original-string and equal-copy controls. Three warmups precede nine rounds.
Source construction, parsing, binding, initial computation and getter evaluation
are excluded; the measured operation is retrieving a cached result.

Paired before/after/after/before runs on Node 24.18.0 / Ryzen 7 9800X3D:

| 1 MB source, 1,000 warm hits | Before median ms | After median ms |
| --- | --- | --- |
| Equal copy, local values | 16.6019–16.7490 | 0.0448–0.0496 |
| Equal copy, unreachable statements | 16.9956–17.7089 | 0.0423–0.0432 |
| Equal copy, function call | 16.6462–16.7819 | 0.0726–0.0783 |
| Same string, local values | 0.0118–0.0122 | 0.0124–0.0218 |
| Same string, unreachable statements | 0.0123–0.0125 | 0.0123–0.0198 |
| Same string, function call | 0.0398 | 0.0400–0.0401 |

The same-string controls are mixed, with nanosecond-scale overhead in some runs.
The first equal-copy comparison is still required, and a fresh string on every
lookup can still require it. These component measurements do not establish
complete diagnostics or editor latency improvements. Private effect/declaration
readers and the array-return-shape cache remain separate audit candidates.

## Validation

Type checking and 30 focused tests pass, covering activity/symbol/source
invalidation, shared readers, reachability and function results. The frozen
baseline differential includes 8,256 corpus sources and 27 generated cases:
13,452 diagnostics, no internal errors, identical complete inference facts and
diagnostic arrays. Independent source A/B/A checks verify local values, dead
statement counts and function literals for LF/CRLF/CR.

The full suite passes: 735 files and 14,523 tests (7 files and 33 tests skipped).
