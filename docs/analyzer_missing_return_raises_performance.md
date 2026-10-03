# Missing-return inspection after a raise

The missing-return helper returns true once it recognizes an active Err.Raise or Error statement. Previously its callback still sliced and trimmed every subsequent statement, checked declaration and raise regexes, and incremented an executable count that could no longer change the result. The callback now returns before that work once the raise flag is true, matching the adjacent return-assignment helper's existing monotonic-result guard.

Structural traversal and conditional activity checks continue. Recognition of raises, declarations, comments, interface stubs, return assignments, nested/one-line constructs, and Function/Property Get diagnostics is unchanged. This does not alter the existing textual raise predicate or introduce control-flow proofs.

## Deterministic cost

Actual leaf-span getters on parsed nodes counted all public checkMissingReturnAssignments reads with 1,000 trailing Beep statements. Counters were outside timed runs; production has no instrumentation.

| Fixture | Before span reads | After span reads | Missing-return diagnostics |
| --- | ---: | ---: | ---: |
| Early Err.Raise | 3,003 | 1,003 | 0 |
| Early Error | 3,003 | 1,003 | 0 |
| Late Err.Raise | 3,003 | 3,003 | 0 |
| No raise | 3,000 | 3,000 | 1 |
| Early actual return | 1 | 1 | 0 |

The guard eliminates two span reads and a source slice per trailing statement in the raise-recognition callback. The previous return-assignment scan still inspects the body. Traversal remains linear; this change does not make the public rule constant-time after a raise.

## Reproduction and timings

From the repository root with installed dependencies:

```powershell
node scripts/benchmark-missing-return-raises.mjs --baseline=30222875
node scripts/benchmark-missing-return-raises.mjs
```

The baseline replaces only assignments.ts from that commit in the current harness. A typed Function contains the listed number of Beep statements, with a raise at the start/end, no raise, Error at the start, or an actual return assignment at the start. The benchmark checks one missing-return diagnostic only for the no-raise fixture. Parsing/binding are outside timing, the parsed body is reused, and each sample receives a fresh bound root with an identity assertion. These measurements cover the public missing-return rule, not complete analyzer/editor latency. Three warmups precede 15 samples. Before and after ran sequentially after all audit tests and differential probes completed.

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor; measured locally on 2026-10-03. Times in milliseconds.

| Statements / mode | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: |
| 10-early-raise | 0.025 | 0.021 | 0.034 | 0.035 |
| 10-late-raise | 0.019 | 0.019 | 0.038 | 0.024 |
| 10-no-raise | 0.018 | 0.016 | 0.088 | 0.067 |
| 10-early-error | 0.014 | 0.011 | 0.086 | 0.038 |
| 10-early-return | 0.002 | 0.002 | 0.005 | 0.006 |
| 100-early-raise | 0.087 | 0.082 | 0.245 | 0.387 |
| 100-late-raise | 0.048 | 0.049 | 0.082 | 0.257 |
| 100-no-raise | 0.042 | 0.042 | 0.059 | 0.186 |
| 100-early-error | 0.043 | 0.040 | 0.050 | 0.154 |
| 100-early-return | 0.002 | 0.002 | 0.008 | 0.009 |
| 1000-early-raise | 0.434 | 0.371 | 0.692 | 0.596 |
| 1000-late-raise | 0.364 | 0.388 | 0.479 | 0.464 |
| 1000-no-raise | 0.278 | 0.272 | 0.577 | 0.529 |
| 1000-early-error | 0.281 | 0.245 | 0.498 | 0.260 |
| 1000-early-return | 0.005 | 0.005 | 0.005 | 0.010 |
| 10000-early-raise | 2.794 | 2.499 | 2.979 | 4.046 |
| 10000-late-raise | 2.834 | 2.700 | 3.380 | 3.510 |
| 10000-no-raise | 2.778 | 2.962 | 3.436 | 3.675 |
| 10000-early-error | 3.269 | 3.288 | 4.187 | 3.865 |
| 10000-early-return | 0.039 | 0.039 | 0.063 | 0.082 |

Whole-rule differences are modest and mixed, including regressions in controls and some early-raise p95 results. The 10,000-statement early Err.Raise median changes from 2.794 to 2.499 ms, while early Error is essentially unchanged at 3.269 versus 3.288 ms. The separate return-assignment scan and traversal remain dominant. The deterministic counters establish eliminated callback work; these timings do not establish a general end-to-end latency win.

## Validation

- Type check passed.
- Focused assignment tests: 176 passed.
- Full suite: 618 files, 12,765 tests passed, 13 skipped.
- Eleven regressions cover actual read bounds, late/no-raise/actual-return controls, active/inactive conditional raises, continued activity checks after a raise, interface stubs, Function/Property Get behavior, separate procedures and existing nested/one-line recognition. Watched parser properties are restored after each cost probe.
- 1,500 differential cases preserve complete analyzer and direct missing-return outputs across early/late/no raises, labels/literal text, declarations, interface stubs, conditional builds, loops/branches, scalar/object returns, ByRef calls and multiple procedures.

Generated differential and counting probes remain audit scratch material; the timing benchmark and regression tests are committed.
