# Batched array-shape invalidation

The array-bound flow model retains immutable maps for earlier statements and branch entries. Its redimShapesAt forget callback copied the entire current map for each removed tracked name. Targeted labels, GoSub, and blocks touching many tracked arrays therefore added quadratic copy work within a single invalidation.

The callback now copies lazily on its first actual removal, then deletes the other names from that new map. Each callback invocation still creates an independent state when needed. Earlier returned snapshots and restored branch entries remain unchanged; an invocation containing only untracked names retains the shared map.

## Reproduction

```powershell
node scripts/benchmark-array-shape-forget.mjs --baseline=4efed8bb --rounds=15
node scripts/benchmark-array-shape-forget.mjs --rounds=15
```

The benchmark times public redimShapesAt with parsing and binding excluded, three warmups, 15 samples and warmed token caches. A procedure declares and ReDims each dynamic array separately, reads an ordinary statement, then jumps to a label that clears the tracked facts. The control omits the jump and label. The harness verifies earlier snapshots remain full and that the jump target clears later shapes.

Node 24.18.0, AMD Ryzen 7 9800X3D; median milliseconds from sequential runs after validation completed:

| Arrays | Targeted label before | After | No-label control before | After |
| ---: | ---: | ---: | ---: | ---: |
| 100 | 0.438 | 0.306 | 0.261 | 0.250 |
| 1,000 | 27.650 | 13.704 | 14.340 | 14.294 |
| 3,000 | 366.471 | 137.150 | 181.932 | 177.011 |

These are isolated-function stress measurements, not editor latency. Establishing the individual shapes still copies increasingly large maps and is included in these timings; this change makes each invalidation linear, not the entire function. Absolute measurements vary with load and allocation/GC behavior. Initial runs measured the largest targeted-label fixture at 341.089 ms before and 306.332 ms after while its unchanged control moved from 178.627 to 282.669 ms. The final run followed full-suite completion and ran baseline then optimized sequentially.

## Validation

- Type checking and the full suite passed: 593 files, 12,501 tests passed, 13 skipped.
- 86 targeted array/setup tests passed.
- 1,500 generated fixtures matched every returned shape map, entry order and shared-map identity against baseline 4efed8bb. Fixtures include partial and duplicate erasures, jump targets, GoSub, loops, conditional compilation, branch restoration, untracked names and later ReDims.
- For 100 arrays, shape entries copied fell from 10,000 to 5,050 overall. Establishing shapes accounts for 4,950 in both versions; clearing them falls from 5,050 to 100. An operation-budget regression bounds the total independently of timings.
- Tests preserve earlier snapshots, untouched names and branch-entry states, and require no-op invalidation to retain sharing.
