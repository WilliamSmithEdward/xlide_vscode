# Host collection Item query performance

Baseline: published main `ad4aa50ac171ca611677bbcc7323a64d909162a6`.

`checkRuntimeMemberNotFound` inferred the host collection for every `For Each`, then searched its members for the first exact-case `Item`. With L loops and M members this repeated L × M name reads when Item was last or absent. A lazy query in the existing collection query bundle now searches once per model and qualified collection during the public invocation. It caches the member object or absence; return types are still read from the member. No cache survives the invocation.

The search preserves exact case, first matching member, event exclusion through `getHostMembers`, unknown/unqualified return handling, and the Range special case. A long-lived case-folded member index cannot replace this search: an earlier `item` is distinct from `Item`, and mutable member names may change between invocations.

## Reproduction

```powershell
node scripts/benchmark-host-collection-item-queries.mjs --baseline=ad4aa50ac171ca611677bbcc7323a64d909162a6
node scripts/benchmark-host-collection-item-queries.mjs
```

Run sequentially in baseline/fixed/fixed/baseline order. Node 24.18.0, AMD Ryzen 7 9800X3D, esbuild 0.28.2. Three warmups and nine timed runs per fixture. Each range below spans the medians from two runs. Public rule timings exclude parse/symbol setup; complete-module timings include it. The host member index is warm. Member getter and Array.find instrumentation runs only outside timing and is restored before timing. Synthetic models have max(1, loop count) members; canonical uses the bundled Excel model.

## Results

| Model | Loops | Scope | Baseline ms | Fixed ms |
| --- | ---: | --- | ---: | ---: |
| last | 0 | public-rule | 0.0670–0.0708 | 0.0596–0.0607 |
| last | 0 | complete-module-diagnostics | 0.7098–0.8289 | 0.6867–0.7850 |
| last | 1 | public-rule | 0.1087–0.1382 | 0.0874–0.0913 |
| last | 1 | complete-module-diagnostics | 1.2143–1.3773 | 1.2688–1.3107 |
| last | 100 | public-rule | 0.7059–0.8152 | 0.6592–0.6928 |
| last | 100 | complete-module-diagnostics | 9.2847–10.8092 | 9.2714–9.7661 |
| last | 1000 | public-rule | 9.1686–9.6677 | 3.0943–3.1287 |
| last | 1000 | complete-module-diagnostics | 87.3235–90.4196 | 81.6384–82.4618 |
| absent | 0 | public-rule | 0.0404–0.0407 | 0.0397–0.0398 |
| absent | 0 | complete-module-diagnostics | 0.3648–0.3700 | 0.3500–0.3662 |
| absent | 1 | public-rule | 0.0372–0.0380 | 0.0371–0.0379 |
| absent | 1 | complete-module-diagnostics | 0.5752–0.6266 | 0.5641–0.5831 |
| absent | 100 | public-rule | 0.3407–0.3487 | 0.2703–0.2793 |
| absent | 100 | complete-module-diagnostics | 6.3140–6.5176 | 6.0702–6.4206 |
| absent | 1000 | public-rule | 9.0872–10.6264 | 2.5268–2.5630 |
| absent | 1000 | complete-module-diagnostics | 96.7149–108.7910 | 78.0647–78.2456 |
| first | 0 | public-rule | 0.0352–0.0586 | 0.0590–0.0590 |
| first | 0 | complete-module-diagnostics | 0.3704–0.4839 | 0.4879–0.4979 |
| first | 1 | public-rule | 0.0656–0.0658 | 0.0374–0.0408 |
| first | 1 | complete-module-diagnostics | 0.8000–0.8362 | 0.5745–0.5952 |
| first | 100 | public-rule | 0.5058–0.5396 | 0.3024–0.5116 |
| first | 100 | complete-module-diagnostics | 10.4236–14.3298 | 8.4672–10.5416 |
| first | 1000 | public-rule | 3.6535–4.9937 | 3.3045–3.3098 |
| first | 1000 | complete-module-diagnostics | 117.7054–122.8007 | 107.5657–114.1083 |
| canonical | 0 | public-rule | 0.0586–0.0622 | 0.0583–0.0612 |
| canonical | 0 | complete-module-diagnostics | 0.5174–0.5633 | 0.4456–0.5269 |
| canonical | 1 | public-rule | 0.0649–0.0716 | 0.0644–0.0768 |
| canonical | 1 | complete-module-diagnostics | 0.7930–0.8539 | 0.5685–0.6348 |
| canonical | 100 | public-rule | 0.3081–0.5591 | 0.5400–0.5421 |
| canonical | 100 | complete-module-diagnostics | 10.1627–10.3722 | 7.0647–7.3183 |
| canonical | 1000 | public-rule | 3.6166–5.0120 | 4.8918–4.9326 |
| canonical | 1000 | complete-module-diagnostics | 109.4276–118.4423 | 89.8942–105.8849 |

At 1,000 loops × 1,000 members, last/absent Item name reads fall from 1,000,000 to 1,000. With Item first, they fall from 1,000 to 1. Zero-loop fixtures do not allocate the cache. Synthetic larger models improve in both scopes. Small cases and canonical public-rule results overlap or are slower in fixed runs; this is not a universal speedup. These measurements do not establish cold-start, heap, native-editor, tree, or tab-switch latency.

## Validation

The 17 dedicated tests cover six work bounds and 11 semantic controls. The baseline fails all six bounds and passes the semantic controls; the repair passes all 17. Cases include wrong case before exact case, only wrong case, duplicate Item, events, scalar/unknown returns, Range, separate models, and names/return types mutated between invocations. Benchmark assertions independently require exact public findings and corresponding full diagnostics, no internal errors, and identical complete output digests across all four trials.

Baseline differential comparisons found zero differences across 2,000 generated host-model modules, 20,000 runtime-member public queries, 2,000 generated metadata modules, 10,000 generated call modules, 2,000 fragment modules, 10,000 four-API private query bundles, 10,000 Is condition queries with 2,000 corresponding full modules, and 24,768 corpus runs (8,256 sources under LF/CRLF/CR). These are regression checks for this repair, not proof of complete analyzer or repository coverage.

Final type checking passed. Full-suite result: Test Files  806 passed | 7 skipped (813); Tests  15903 passed | 33 skipped (15936); Duration  85.58s (transform 59.90s, setup 0ms, import 584.55s, tests 470.51s, environment 76ms).
