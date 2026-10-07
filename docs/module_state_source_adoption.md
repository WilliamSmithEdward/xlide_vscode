# Module-state cache source adoption

`untouchedModuleVariables` and `untouchedModuleVariablesIn` share a weak cache
keyed by bound module symbols, source text, and project-write-set identity.
A matching hit previously retained the original source string. Equal text in
another string could therefore require a full text comparison on every hit.
The hit now adopts the caller's source after all existing keys match. Facts,
procedure projections, identity, invalidation and cache cardinality are unchanged.

## Reproduction

Run `node scripts/benchmark-module-state-source.mjs --rounds=9`, or add
`--baseline=73023d76e7e0f240abd0f77ea18f97fe320deebc` for the prior implementation.
The script measures 1,000 warm hits for both public consumers at approximately
1 KB, 100 KB and 1 MB, with same-string and equal-copy controls. Every answer is
checked against an independently expected variable and the primed result identity.
Construction, parsing, symbol binding and priming are excluded. Three warmup
rounds precede nine measured rounds; each round primes fresh symbols and source.

On Node 24.18.0 / Ryzen 7 9800X3D, before/after/after/before paired runs gave:

| 1 MB source / 1,000 hits | Before median ms | After median ms |
| --- | --- | --- |
| Equal copy, module scope | 16.7928–16.8928 | 0.0641–0.0657 |
| Equal copy, procedure scope | 16.8188–16.8734 | 0.0687–0.0738 |
| Same string, module scope | 0.0201 | 0.0200–0.0205 |
| Same string, procedure scope | 0.0339–0.0374 | 0.0242–0.0246 |

These measurements concern warm cache hits, not complete diagnostics or editor
latency. The first equal-copy comparison remains necessary. Continuously supplying
a new string on every call may continue to incur that comparison.

## Validation

Type checking and the 27 existing module-state tests pass, including changed
source, project-write changes and procedure-local shadowing. The accompanying
audit also compares complete facts and diagnostic arrays against the baseline
with frozen syntax trees and lexer tokens: 8,256 corpus sources plus 24 generated
cases, 13,452 diagnostics, no internal errors and no differences.

The full suite passes: 733 files and 14,503 tests (7 files and 33 tests skipped).
