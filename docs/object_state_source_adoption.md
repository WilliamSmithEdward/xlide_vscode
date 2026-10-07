# Object-state cache source adoption

Module object facts are retained by parsed module, source, conditional activity
and member completion context. Procedure walks are weakly owned by bound symbols
and procedure, and also check module, source, activity and member context.
Equal-source hits previously kept the original string and could compare the
whole module on every later query. Fully matching hits now adopt the caller
string; all fact/findings objects, context keys, ownership and invalidation
rules are unchanged.

## Reproduction and scope

Run `node scripts/benchmark-object-state-source.mjs`, or add
`--baseline=2974c8fa9593f1199d5680029e17905d73756f5d` for the prior implementation.
The private queries are exposed only in the temporary benchmark bundle.
Both queries are measured at approximately 1 KB, 100 KB and 1 MB, using original
strings and equal copies. Three warmups precede nine measured rounds; each round
constructs fresh source, parsed module, symbols and member context. Independent
complete assertions check Nothing-returning functions, first object-member reads,
Excel range methods, Let states and findings with absolute spans. Result identity
is checked on every hit. Construction, parsing, binding and priming are excluded.

Before/after/after/before paired runs on Node 24.18.0 / Ryzen 7 9800X3D:

| 1 MB source, 1,000 warm hits | Before median ms | After median ms |
| --- | --- | --- |
| Equal copy, module facts | 16.7641–17.2450 | 0.0553 |
| Equal copy, procedure walk | 16.8454–16.8898 | 0.0557–0.0562 |
| Same string, module facts | 0.0206–0.0227 | 0.0203–0.0209 |
| Same string, procedure walk | 0.0214–0.0218 | 0.0220–0.0224 |

Same-string controls are mixed; procedure controls add about 0.2–1 ns per hit.
The first equal-copy comparison is still necessary, and a fresh string on every
call may continue to incur it. These are cache-query measurements, not complete
object-state analysis, full diagnostics or editor latency.

## Validation

The linked PR records final type/focused/full-suite and frozen baseline counts.
Complete module facts, procedure findings/Let states and diagnostic arrays are
compared across the syntax corpus and generated line endings, conditional object
assignments, declared object types and auto-instantiation cases. Member contexts
switch between Excel and Word; activity and source replacement are also checked.
Independent source A/B/A assertions verify Nothing-returning functions and set/
unset Let states while preserving fact identity on equal copies.

Type checking and 210 focused tests pass. The frozen baseline differential
includes 8,256 corpus sources and 36 generated cases: 13,482 diagnostics,
no internal errors and identical complete module/procedure facts and diagnostics.

The full suite passes: 737 files and 14,560 tests (7 files and 33 tests skipped).
