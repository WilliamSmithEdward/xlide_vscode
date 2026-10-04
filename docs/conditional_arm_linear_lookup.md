# Linear conditional arm comparison

Conditional activity tracking gives each newly opened `#If` chain a unique increasing ID. Every immutable parent path therefore descends by chain ID. The previous `armsDiverge` implementation cross-scanned the two paths to find their innermost shared chain. Disjoint deep chains required quadratic work on every comparison made by duplicate-declaration and control-flow checks.

The repair merges the two descending paths, advancing the larger ID until a shared chain is found. It preserves the existing innermost shared-chain decision and returns immediately for an identical branch object. It adds no map, retained cache or branch metadata.

Validation against `60b20fa0bfe5326a12c3ef0b26beb0fd1b948da0`:

- Types pass; 62 focused tests pass, including 16 new deep branch/diagnostic controls. Those 16 also pass on baseline: this is a performance repair preserving behavior.
- Full suite: 770 files / 15,035 tests pass, 7 files / 33 tests skipped, 77.16 seconds.
- 200 independently generated conditional schedules produce 292,783 span pairs. Baseline and repaired results both match an independent path model that checks whether any shared chain selects different arms.
- 8,256 oracle sources across LF, CRLF and CR produce 24,768 complete `analyzeModule` diagnostic and internal-failure comparisons, with zero differences.

Reproduce the benchmark:

```powershell
node scripts/benchmark-conditional-arm-lookup.mjs --baseline=60b20fa0bfe5326a12c3ef0b26beb0fd1b948da0
node scripts/benchmark-conditional-arm-lookup.mjs
```

Node v24.18.0, AMD Ryzen 7 9800X3D. ABBA order, two baseline and two repaired runs, three warmups and nine measured rounds per scenario. Tracker construction/source preparation and module AST warming stay outside the clock. Tracker scenarios average 100 queries per round. Full-diagnostic scenarios analyze two nested conditional chains containing 33 same-name declarations and independently check all 32 diagnostic objects, their spans and internal failures on every call.

| Scenario | Baseline median ms | Repaired median ms |
| --- | ---: | ---: |
| 1-level independent tracker query | 0.00006–0.00022 | 0.00012–0.00023 |
| 100-level independent tracker query | 0.00843–0.00947 | 0.00023–0.00038 |
| 1,000-level independent tracker query | 1.51673–1.52598 | 0.00202–0.00206 |
| 1,000-level alternative-arm query | 1.49633–1.62554 | 0.00388–0.00415 |
| 1,000-level shared-parent query | 1.50274–1.64896 | 0.00374–0.00471 |
| Complete diagnostics, 1 level | 1.3724–1.4850 | 0.9824–1.0322 |
| Complete diagnostics, 100 levels | 3.7455–4.1619 | 3.7925–4.0503 |
| Complete diagnostics, 1,000 levels | 86.7047–94.2622 | 35.4354–36.3419 |

A separate export-only temporary bundle exposes the existing private comparison helper to observable getter-backed immutable paths. For one independent comparison, chain-property reads change from 2 / 20,000 / 2,000,000 at depths 1 / 100 / 1,000 to 4 / 400 / 4,000. The harness verifies those work bounds without a timing threshold. Production code has no measurement instrumentation.

This improves the measured deep-conditional path. Small-input costs are mixed, and the deep case is a stress workload. There is no cold-parser, retained-heap, worker-roundtrip or renderer-latency claim.
