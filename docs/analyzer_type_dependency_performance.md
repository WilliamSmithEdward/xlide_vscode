# Cross-module Type dependency queue audit

## Finding

The declaration-order rule used a breadth-first queue for user-defined Types. It marked a Type as seen only after removing it with Array.shift(), so converging edges queued the same Type repeatedly. Each shift also removed the array's first element. Broad shared dependency graphs accumulated large queues despite visiting each Type's fields only once.

The fix marks canonical Type surfaces when enqueued and uses an advancing queue cursor. The unambiguous-name table supplies one surface object for each resolved name, so object identity preserves the previous case-insensitive module/name deduplication. BFS order, current-module exclusions, ambiguous-name handling, and the reported first external Type are preserved. State remains local to each traversal; later rule passes see changed project metadata.

## Measurements

Run node scripts/benchmark-type-dependency-queue.mjs --rounds=15 in the checkout being measured. Baseline: f8764ea0. Node v24.18.0 on AMD Ryzen 7 9800X3D; three warmups, 15 samples, no concurrent test process.

Each fixture has an entry Type pointing to W branch Types, each branch pointing to the same W leaf Types. Leaves contain Long fields; the cyclic case instead ends with a leaf referring to the source Type. The source contains one Type field, and parsing and metadata construction are outside the timed declaration-order rule pass.

| Fixture | Before median / p95 ms | After median / p95 ms |
| --- | ---: | ---: |
| W=25, acyclic | 0.150 / 0.321 | 0.067 / 0.092 |
| W=100, acyclic | 1.216 / 3.957 | 0.524 / 1.032 |
| W=300, acyclic | 252.420 / 1843.912 | 3.165 / 3.927 |
| W=300, cyclic | 6.430 / 10.545 | 4.457 / 6.330 |

The broad acyclic fixture is a synthetic stress case with 601 Types and 90,000 branch-to-leaf edges; the original large-queue timings were particularly variable. These numbers do not measure end-to-end analyzer latency or imply that typical projects improve by 80x. The traversal still examines graph edges, and each source Type field starts its own traversal.

## Validation

The queue-work regression counts Type surfaces appended while two source fields traverse a 121-Type shared graph: the original code queues 7,320 surfaces, exceeding a bound of 484; the optimized code stays within that node-count bound. Other regressions cover cycles with shared descendants and back edges, per-field diagnostics and spans, case-insensitive names, ambiguous names, current-module exclusions, and changed metadata on a subsequent pass.

Existing declaration-order tests pass. An additional deterministic comparison of 500 generated graphs produced identical diagnostic kinds, messages, and spans against the original implementation.
