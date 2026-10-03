# Numeric overflow folding audit

## Finding

TypedFolder.logical scanned each expression separately for Imp, Eqv, Xor, Or, and And. Arithmetic expressions with none of these operators paid for five complete token passes on every fold.

The replacement scans once, choosing the last top-level operator with the lowest precedence. Equal-precedence operators still split at the last occurrence, preserving left associativity. Parenthesis depth, the exclusion of token zero, keyword matching, numeric conversion, short-circuiting on unknown or overflowing operands, and diagnostic construction remain unchanged.

## Measurements

Run node scripts/benchmark-overflow-folding.mjs --rounds=15 in the target checkout. Baseline: f8764ea0. Node v24.18.0 on AMD Ryzen 7 9800X3D; three warmups and 15 samples, with no concurrent tests. Parsing and symbol construction are excluded; the public checkOverflow rule is timed with warm analyzer caches and produces zero diagnostics for every fixture.

| Fixture | Before median / p95 ms | After median / p95 ms |
| --- | ---: | ---: |
| 300 procedures, 12 arithmetic assignments each | 10.175 / 11.749 | 9.479 / 10.841 |
| One procedure, 10,000 arithmetic assignments | 29.814 / 31.807 | 27.012 / 28.101 |
| 1,000 assignments, 100 arithmetic operands each | 39.626 / 44.899 | 32.635 / 35.181 |
| 1,000 mixed logical assignments | 5.950 / 8.721 | 4.803 / 5.917 |

These are isolated overflow-rule measurements, not whole-analyzer or editor latency. Recursive operand folding and token slicing remain; this change does not make arbitrary deeply nested or long logical expressions linear-time.

## Validation

The existing overflow, logical conversion, nested argument, condition, and branch suites pass (99 tests). Eight additional Byte Const regressions check mixed precedence, left-associative Imp, parenthesized operands, and Not binding below arithmetic. An additional deterministic comparison of 500 mixed numeric, Boolean, string, Currency, and LongLong expressions matches original diagnostic kinds, messages, and spans exactly.
