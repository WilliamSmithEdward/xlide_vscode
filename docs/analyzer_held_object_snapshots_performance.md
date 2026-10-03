# Unchanged held-object snapshots

heldObjectsAt previously deep-copied the object-class map, collection map and every collection item array for each observed statement and block entry. Most ordinary statements do not change those facts. The returned API exposes read-only facts, and its consumers read values rather than relying on distinct snapshot identities.

Cache the copied snapshot until tracked state changes. Invalidate it on class/item deletions, assignments, collection additions and changing member calls. Branch restoration still clones maps and item arrays into mutable internal state, while reusing the immutable saved snapshot for observations until another mutation. Statement and block snapshots now use one helper instead of duplicated implementations.

Unchanged snapshots intentionally share identity; their contents remain immutable from the analyzer's point of view. Actual observed mutations still produce independent deep snapshots. The cache is local to one helper invocation and creates no cross-procedure or cross-activity state.

## Reproduction

```powershell
node scripts/benchmark-held-object-snapshots.mjs --baseline=59242ecf --rounds=15
node scripts/benchmark-held-object-snapshots.mjs --rounds=15
```

The benchmark times public heldObjectsAt after parsing and binding, with warmed token caches, three warmups and 15 samples. The object fixture declares 100 auto-instantiated objects and runs ordinary statements. The collection fixture adds 100 objects then repeatedly reads Count. The empty-state control runs ordinary statements with no held facts. Each fixture asserts the final class-map size and, for the collection, item count.

Node 24.18.0, AMD Ryzen 7 9800X3D; median milliseconds:

| Unchanged statements | 100 objects before | After | 100 collection items before | After | Empty control before | After |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 100 | 0.624 | 0.075 | 0.217 | 0.093 | 0.050 | 0.032 |
| 1,000 | 7.325 | 0.320 | 0.839 | 0.238 | 0.195 | 0.139 |
| 3,000 | 17.613 | 0.420 | 2.699 | 0.806 | 0.542 | 0.423 |

These isolated-helper stress fixtures exclude parsing/binding and do not represent editor latency. Absolute timings vary with load. Collection setup is timed and still creates distinct snapshots for actual additions; retained changing facts legitimately require copies.

## Validation

- Type checking passed; full suite: 599 files, 12,544 tests passed, 13 skipped.
- 84 targeted tests passed.
- 3,000 generated fixtures matched complete held-object state facts against baseline 59242ecf, including class/item entries and order. Another 500 comparisons matched complete analyzer diagnostics, covering consumers of these facts.
- Copied object-map entries for 100 objects and 200 unchanged statements fell from 20,000 to 100. The regression test bounds that work independently of timing and verifies snapshot sharing.
- Six regressions cover unchanged maps, historical collection contents, object and literal additions, Before insertion, branch restoration, aliasing, changing members, collection replacement and activity changes on reused parsed nodes.
