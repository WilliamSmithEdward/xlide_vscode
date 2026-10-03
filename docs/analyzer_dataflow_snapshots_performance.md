# Bounded reaching-value snapshots

Cold dataflow walks previously copied every reaching fact after each assignment. Unchanged module constants amplified the cost: 3,000 constants and 1,000 known-string array reads copied 3,007,001 entries into 2,002 Maps. The updated walker copies 7,001 entries into 2,003 Maps, producing the same 1,000 diagnostics. These counts come from temporary instrumentation of the walker and snapshot helper, outside timed runs; production has no counters.

Large states now share a detached base with at most 32 overrides. Lookup has one base level; adding a 33rd distinct override materializes a native Map. States smaller than 64 entries use native Map copies. Single-value assignment, call-effect, array-element and loop-counter updates use the helper. Invalidation and branch joins still materialize states as needed. Historical states retain exact token references and spans, size, lookup, iteration order and forEach behavior. The first shared base is copied once to protect against mutation of a caller-owned initial Map.

## Reproduction

Run from the repository root with installed dependencies:

```powershell
node scripts/benchmark-dataflow-snapshots.mjs --baseline=55a6dad50571baa12b47d1d7bb79446ff1b0c5be
node scripts/benchmark-dataflow-snapshots.mjs
```

The baseline replaces only straightLineValues.ts from that commit, using the same current harness and other analyzer modules. Each fixture has the listed number of module constants and 1,000 assignments. Cold fixtures read a Variant array containing strings into a Long; conditional fixtures select Option Base 1 with VBA7. Scalar controls assign a number without requesting a dataflow walk.

Parsing, binding and tracker construction are outside timing. Every cold sample appends a unique comment and asserts that the parsed procedure body differs from the previous sample: parsing the same source alone is insufficient because parseModule caches bodies. Fresh bound roots are asserted too. Warm fixtures intentionally reuse the body and cached walk. Each measured invocation checks its diagnostic count. Three warmups precede 15 samples. Before and after ran sequentially with no concurrent audit tests or benchmarks.

## Results

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor. Times in milliseconds, measured locally on 2026-10-03.

| Constants / mode | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: |
| 10-cold-default | 11.180 | 10.960 | 19.138 | 15.328 |
| 10-cold-conditional | 9.993 | 9.666 | 10.713 | 13.108 |
| 10-warm-default | 3.337 | 3.260 | 4.382 | 4.514 |
| 10-scalar-control | 1.643 | 1.599 | 2.488 | 2.331 |
| 100-cold-default | 12.199 | 9.405 | 18.462 | 10.146 |
| 100-cold-conditional | 12.599 | 10.337 | 13.004 | 15.361 |
| 100-warm-default | 3.463 | 3.548 | 4.493 | 4.305 |
| 100-scalar-control | 1.492 | 1.499 | 2.876 | 1.990 |
| 1000-cold-default | 43.646 | 11.355 | 47.258 | 11.834 |
| 1000-cold-conditional | 51.304 | 19.461 | 58.183 | 25.611 |
| 1000-warm-default | 5.064 | 5.016 | 6.942 | 6.032 |
| 1000-scalar-control | 1.830 | 1.805 | 3.714 | 2.733 |
| 3000-cold-default | 134.851 | 16.736 | 183.552 | 26.677 |
| 3000-cold-conditional | 159.304 | 42.140 | 278.685 | 44.581 |
| 3000-warm-default | 8.885 | 9.130 | 13.587 | 13.617 |
| 3000-scalar-control | 2.551 | 2.584 | 4.953 | 4.984 |

The 3,000-constant cold default fixture improves about 8.1 times, and its conditional counterpart about 3.8 times. Warm walks already bypass the repeated copies. Small controls have mixed differences and some p95 regressions; these measurements do not establish a universal latency improvement. Conditional evaluation and other rule work remain in the timings. Distinct writes periodically compact the state, and iteration or invalidation still visits the base; this change targets repeated writes with many unaffected facts.

## Validation

- Type check passed.
- Full suite: 615 files, 12,740 tests passed, 13 skipped.
- Eight snapshot tests cover threshold boundaries, compaction, retained historical states, mutable initial maps, lookup, ordered iteration, deletion/reinsertion, forEach and a deterministic base-read cost bound.
- 1,000 generated differential cases compare complete analyzer outputs and ordered dataflow snapshots, including exact tokens/spans, exit states and unreachable spans. Fixtures exercise branches, loops, jumps, ByRef effects, declared facts, object/array sentinels and overwritten constants.

The differential harness and instrumentation remain audit scratch material; the reproducible timing benchmark and regression tests are committed.
