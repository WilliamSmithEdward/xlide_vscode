# Returned-array shape cache source adoption

`functionReturnShapes` retains one fact record per weakly owned parsed module.
Source text, conditional activity identity and `Option Base` must all match.
An equal-source hit previously retained the original string, potentially
comparing the whole module on every subsequent query. A fully matching hit now
adopts the caller's string. Facts, result identity, ownership and invalidation
are unchanged.

## Reproduction

Run `node scripts/benchmark-array-return-source.mjs`, or add
`--baseline=8142e5eac7c5e5c897faf3ecf3fb886db64a5891` for the old implementation.
The private query is exposed only in the benchmark's temporary bundle, preserving
the production API. Both `Option Base` values are measured at approximately 1 KB,
100 KB and 1 MB with same-string and equal-copy controls. Three warmups precede
nine rounds; each round uses a fresh source and parsed module. Complete bounds,
values, name and origin are independently asserted before timing. Cached result
identity is checked on every hit. Construction, parsing and priming are excluded.

Before/after/after/before paired runs on Node 24.18.0 / Ryzen 7 9800X3D:

| 1 MB source, 1,000 warm hits | Before median ms | After median ms |
| --- | --- | --- |
| Equal copy, base 0 | 16.6631–16.8419 | 0.0545–0.0564 |
| Equal copy, base 1 | 16.6455–16.6608 | 0.0540–0.0568 |
| Same string, base 0 | 0.0179–0.0297 | 0.0168 |
| Same string, base 1 | 0.0167–0.0168 | 0.0169 |

The base-1 same-string control adds roughly 0.1–0.2 ns per query in these runs.
The first equal-copy comparison remains required; continuously supplying new
strings may continue to incur it. These component measurements do not establish
whole-editor or full-diagnostics latency improvements.

## Validation

The linked PR records the final type, focused/full suite and frozen baseline
comparison counts. Facts and diagnostic arrays are compared across the syntax
corpus and generated LF/CRLF/CR, conditional declarations, qualified/nested Array
constructors and fixed-array returns. Independent source A/B/A and base 0/1/0
controls verify complete bounds and element values on a retained parsed module.

Type checking and 88 focused tests pass. The frozen differential includes
8,256 corpus sources and 24 generated cases: 13,440 diagnostics, no internal
errors, and identical complete returned-array fact maps and diagnostic arrays.

The full suite passes: 736 files and 14,528 tests (7 files and 33 tests skipped).
