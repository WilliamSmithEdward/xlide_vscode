# Retained array-shape snapshots

redimShapesAt previously copied its entire map on every ReDim, even when no statement or branch retained the current map. Consecutive ReDims therefore paid quadratic setup cost before the first ordinary statement that saw the accumulated shapes. Batched invalidation in PR #792 removed repeated copies during one forget operation but left these setup copies.

The helper now tracks retained maps in a per-invocation WeakSet. Returning a map for a statement or saving it at a branch entry marks it retained. A subsequent mutation copies a retained map once; unobserved intermediate maps can be mutated directly. Restored branch maps remain retained. Earlier statement results keep their contents, order and sharing. The redundant seen alias is removed because shapes itself is the current map.

## Reproduction

```powershell
node scripts/benchmark-array-shape-snapshots.mjs --baseline=0233b9bd --rounds=15
node scripts/benchmark-array-shape-snapshots.mjs --rounds=15
```

The benchmark times public redimShapesAt after parsing and binding, with three warmups, 15 samples and warmed token caches. A procedure separately declares and ReDims each array, then runs two ordinary statements. The label fixture adds a GoTo and its targeted label, which invalidates all shapes. Assertions verify snapshot sizes and that shapes are cleared at the target.

Node 24.18.0, AMD Ryzen 7 9800X3D; median milliseconds:

| Arrays | Targeted label before | After | Setup without label before | After |
| ---: | ---: | ---: | ---: | ---: |
| 100 | 0.439 | 0.221 | 0.343 | 0.137 |
| 1,000 | 35.407 | 1.394 | 14.704 | 0.974 |
| 3,000 | 376.739 | 2.926 | 185.764 | 2.921 |

These isolated-function stress fixtures exclude parsing and do not measure editor latency. The setup fixture isolates work left after PR #792; the label fixture compares against main before that PR and benefits from both reduced setup and reduced invalidation copies. Actual output snapshots can still require large maps: alternating shape changes and observing statements retains distinct states and legitimately requires copies. Absolute timings vary with load and allocation/GC behavior.

## Validation

- Type checking passed; full suite: 594 files, 12,510 tests passed, 13 skipped.
- 85 targeted tests passed.
- 3,000 generated fixtures match all returned map contents, entry order and shared-map identities against main 0233b9bd and the batched-invalidation implementation 4f1a96fe. Fixtures cover empty/nonempty branch entries, erasures, later ReDims, unknown/computed bounds, targeted labels, GoSub, nested blocks, loops, conditional compilation, single-line branches and observed intermediate states.
- Copied entries for 100 arrays plus a targeted label: 10,000 on main; 5,050 after batched invalidation; 100 with retained-map copy-on-write. The operation-budget test requires at most 100.
- Behavioral regressions preserve historical bounds across later known/unknown ReDims, retain empty branch-entry maps before shape creation, and preserve sharing across statements that do not change shapes.
