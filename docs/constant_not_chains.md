# Constant-expression Not chains

A source expression containing 20,000 consecutive `Not` operators threw
`RangeError: Maximum call stack size exceeded` in the shared integer constant
evaluator. Its parentheses and unary-sign paths already checked nesting depth,
but `logical()` recursively called itself for every `Not` without that guard.
The evaluator is used for project-visible constant values and diagnostic facts.

Consume the flat chain in a loop, retaining inversion parity. Evaluate its
operand once at the existing arithmetic precedence, then require a Long operand
even when an even number of operators cancels. The even case still performs a
32-bit conversion so `Not Not -0` remains positive zero. Parentheses and unary
signs retain their existing nesting guards. Flat chains no longer add JavaScript
stack frames; lexing and operator consumption still scale with input length.

## Validation

- Baseline `eaadb089eb9c6e21df58d0c870fc96001bd093ab`: four failing
  long-chain regressions, two passing precedence/nesting controls.
- Seven new tests cover odd/even long chains, lookup count, precedence,
  out-of-Long and missing operands, negative zero, and retained nesting guards.
- Focused suite: 77 tests across four files passed; type checking passed.
- Full suite: 744 files passed, 7 skipped; 14,617 tests passed, 33 skipped
  (79.53 seconds).
- Baseline differential: 6,732 expressions and 51 resolved constant maps matched
  exactly, including complete values and lookup sequences. Inputs include
  malformed expressions, conversions, qualified names, both Long boundaries,
  negative zero, and arithmetic/logical suffixes.

## Measurement

Run `node scripts/benchmark-constant-not-chains.mjs`, optionally with
`--baseline=eaadb089eb9c6e21df58d0c870fc96001bd093ab`.
On Node 24.18.0 / Ryzen 7 9800X3D, separate baseline/fixed/fixed/baseline
processes used three warmups and nine measured rounds. The whole evaluator,
including lexing, was timed. Shorter chains used batches of 100 and report
per-call averages; 20,000/20,001 used individual calls.

| Not count | Baseline median ms | Fixed median ms |
| --- | --- | --- |
| 0 | 0.000663–0.000721 | 0.000672–0.000709 |
| 1 | 0.001116–0.001244 | 0.001183–0.001293 |
| 100 | 0.009978–0.010630 | 0.010096–0.010418 |
| 1,000 | 0.092410–0.098558 | 0.086690–0.088536 |
| 20,000 | Stack overflow in all rounds | 2.4095–2.4210 |
| 20,001 | Stack overflow in all rounds | 3.1933–3.2698 |

Baseline failures have no valid latency measurement. Ordinary cases are similar,
with mixed small changes. The largest fixed 20,000-chain sample was 5.8771 ms;
these results establish stack safety, not whole-analyzer or editor latency.
