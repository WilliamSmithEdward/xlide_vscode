# Analyzer fact-cache audit — 2026-10-03

## Confirmed invalidation failures

The parse cache returns the same immutable AST for the same source on later analysis passes. Compiler flags and symbol snapshots can change while that AST identity stays the same. Several caches assumed that each analysis got newly parsed nodes.

Before-fix regressions reproduced six failures: stale literal function results, UDT array field bounds derived from conditional constants, caller division-by-zero diagnostics, returned-array bounds after a conditional Option Base change, reachability after a conditional constant change, and unstructured-flow classification after a conditional error handler becomes inactive.

The fix scopes module function results and type-field facts by source/activity, returned-array facts by source/activity/Option Base, and procedure flow facts by source/activity. Initial procedure values, unreachable statements and evaluated call results also belong to the current symbol snapshot. All caches retain weak ownership and reuse facts within one matching context. Tests alternate compiler flags in both directions and verify same-context memoization and replacement symbol snapshots.

## Performance controls

No speedup is claimed. Existing `node scripts/benchmark-analyzer.mjs --rounds=15` controls ran before and after the primary diagnostic-cache fix, using three warmups, 15 samples, Node 24.18.0 and AMD Ryzen 7 9800X3D. Baseline source files came from the branch's main base through an isolated esbuild loader; benchmark inputs were identical. Median milliseconds:

| Workload | Before | After |
| --- | ---: | ---: |
| 100 procedures, full warm analysis | 78.399 | 72.273 |
| 100 procedures, body edit | 42.906 | 38.520 |
| 300 procedures, full warm analysis | 178.314 | 162.994 |
| 300 procedures, body edit | 113.700 | 103.395 |
| 6,000 statements, full warm analysis | 196.652 | 184.915 |
| 6,000 statements, body edit | 190.831 | 180.720 |

The measurements varied between runs and do not establish a performance improvement. They found no consistent slowdown in these controls. This table predates the final source/activity guard for unstructured-flow classification. That guard follows the same constant-time lookup policy; its correctness was reproduced separately.

## Audit coverage and remaining work

The rule-level profiler invoked all 150 diagnostic registry entries across many procedures, one large body, array/object branches and qualified declarations. Per-callback timers add overhead, so these profiles guide isolated measurements rather than serving as final performance claims. Assignment checking and overflow folding are the largest rule costs on arithmetic bodies and remain under review.

The cache pass inspected module/procedure fact caches in function results, type fields, array return shapes, type inference, unstructured flow, callee argument lookup, straight-line values, loop counters and receiver bindings. Straight-line values and loop counters already validate source/activity. Callee lookup and receiver binding indexes derive their indexed membership from immutable AST structure; their existing conservative resolution contracts were retained. The analyzer audit remains in progress before the broader repository sweep.
