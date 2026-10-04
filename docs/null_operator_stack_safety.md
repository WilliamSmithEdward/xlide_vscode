# Null operator stack safety

Deep Null propagation previously recursively sliced token arrays. A scalar assignment or argument containing 5,000 `Not` prefixes, split across 20 physical lines (19 continuations; maximum line length 1,009), exhausted the JavaScript stack and lost the expected mismatch diagnostic. The repaired evaluator emits the exact mismatch without an internal failure. This reproduction stays within existing physical line and continuation limits; it does not assert that Excel compiles arbitrarily complex expressions.

The evaluator uses immutable token windows, a query-owned parenthesis index and explicit continuation frames. Unary, Abs, arithmetic and logical paths no longer recurse. Existing one-pair normalization, literal coercions, callback order, short-circuiting and callback error propagation remain unchanged. No persistent cache or arbitrary expression-depth cutoff is introduced. The former runtime dependency through type inference is removed.

## Validation

Type checking and the full suite pass: 773 files, 15,149 tests (7 files / 33 tests skipped), 86.42 seconds. The 48 new tests include deep nested paths, complete scalar assignment/argument diagnostics, deterministic work bounds, frozen inputs and semantic controls. Against baseline `92db624c63439ef44a3e2a8f5e54ddc6a6745237`, 10 of these tests fail and 38 pass. A seeded comparison checks 20,000 valid/malformed expressions with exact callback sequences; 8,256 oracle sources across LF, CRLF and CR produce 24,768 complete diagnostic/failure comparisons with zero differences.

## Work and timing

At 1,000 unary prefixes, observed token references returned by Array.slice fall from 500,500 to zero. At 1,000 nested arithmetic groups, observed rawText reads fall from 7,511,502 to 19,004. These are work counters, not heap-byte measurements.

Run `node scripts/benchmark-null-operator-stack.mjs` and the same command with `--baseline=92db624c63439ef44a3e2a8f5e54ddc6a6745237`. Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor           . Two before/after pairs were executed in ABBA order, sequentially after corpus validation. Each run uses three warmups and nine samples. Tokens and module AST warming are outside the clock; every result and independently specified complete diagnostic object is checked. Instrumented work counters execute before timings, so these timings describe this harness and its JIT state.

| Scope | Shape | Size | Before median ms | After median ms |
|---|---|---:|---:|---:|
| Null-helper | unary | 1 | 0.0005–0.0005 | 0.0004–0.0006 |
| Null-helper | unary | 100 | 0.0058–0.009 | 0.0045–0.0058 |
| Null-helper | unary | 1000 | 0.2614–0.2648 | 0.0304–0.0417 |
| Null-helper | abs | 1 | 0.0003–0.0004 | 0.0006–0.0006 |
| Null-helper | abs | 100 | 0.0713–0.0789 | 0.0105–0.0115 |
| Null-helper | abs | 1000 | 6.532–6.8675 | 0.0931–0.1055 |
| Null-helper | arithmetic | 1 | 0.0006–0.0007 | 0.0012–0.0013 |
| Null-helper | arithmetic | 100 | 0.5526–0.5802 | 0.0559–0.086 |
| Null-helper | arithmetic | 1000 | 44.9124–47.6185 | 0.2358–0.2987 |
| complete-module-diagnostics | unary | 1 | 1.2856–1.501 | 1.3189–1.6594 |
| complete-module-diagnostics | unary | 100 | 2.028–2.0731 | 2.1339–2.2352 |
| complete-module-diagnostics | unary | 1000 | 18.852–18.8529 | 16.3704–18.2993 |

Deep helper calls improve substantially. Tiny Abs/arithmetic cases have additional setup cost; complete-module 100-prefix cases are slower and one-prefix results are mixed. The 1,000-prefix module fixture improves in these runs, but this narrow workload does not establish a general analyzer or editor responsiveness improvement. There is no cold-parser, heap or renderer claim.
