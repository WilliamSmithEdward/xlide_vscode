# Object-read lookup performance

The object-default-value rule previously rebuilt and normalized its host-type Map for every statement/branch span. It also stored late-bound Object names in an array, using includes for each statement token and subsequent read. The type environment and eligible names do not change during one procedure visitor.

Keep the normalized host-type Map and late-bound name Set in that visitor instead. Each procedure still gets its own indexes, preserving local/module shadowing and host selection. Held-object facts remain lazy and statement-specific. The new Set and Map replace arrays of the same eligible names, so retained index space is O(eligible names) for the visitor lifetime; container overhead differs, but nothing is retained beyond the existing visitor closure.

## Reproduce

Run in the checkout containing this script:

```powershell
node scripts/benchmark-object-read-indexes.mjs --baseline=0928416dd2019f5c14dabce0f2b3d5d291be2ae4
node scripts/benchmark-object-read-indexes.mjs
```

The baseline option substitutes only objectValues.ts at build time. Both runs use the same surrounding code. Parse once, create fresh bound symbols before every sample (asserting distinct root identity), then time rule setup and the procedure statement/header visitors. Binding and parsing are outside the timer. Each fixture has the stated number of separate local declarations, one scalar local, and 1,000 unrelated scalar assignments; visits include the declaration statements. Findings must remain empty. Three warmups, 15 measured samples; the script reports median and p95. Run without other benchmarks or tests in parallel.

host-types uses Application locals with the default Excel host; late-bound uses Object locals with Excel. late-bound-no-host uses Object locals with the empty unknown-host model, isolating name membership from host-default map entries. scalar-control uses Long locals, exercising the existing no-object path.

## Results

Measured on 2026-10-03, Node v24.18.0, AMD Ryzen 7 9800X3D 8-Core Processor. Medians in milliseconds:

| Locals and mode | Before | After |
| --- | ---: | ---: |
| 1-host-types | 0.847 | 0.630 |
| 1-late-bound | 0.957 | 0.738 |
| 1-late-bound-no-host | 0.747 | 0.570 |
| 1-scalar-control | 0.400 | 0.284 |
| 10-host-types | 1.119 | 0.461 |
| 10-late-bound | 1.252 | 0.592 |
| 10-late-bound-no-host | 1.207 | 0.595 |
| 10-scalar-control | 0.292 | 0.278 |
| 100-host-types | 5.627 | 0.499 |
| 100-late-bound | 5.651 | 0.714 |
| 100-late-bound-no-host | 5.634 | 0.573 |
| 100-scalar-control | 0.293 | 0.304 |
| 1000-host-types | 52.194 | 0.717 |
| 1000-late-bound | 52.095 | 0.786 |
| 1000-late-bound-no-host | 53.510 | 0.795 |
| 1000-scalar-control | 0.504 | 0.498 |
| 3000-host-types | 171.835 | 1.245 |
| 3000-late-bound | 172.023 | 1.421 |
| 3000-late-bound-no-host | 171.165 | 1.454 |
| 3000-scalar-control | 0.904 | 0.994 |

These are synthetic rule-only stress cases, not end-to-end analyzer latency or typical workbook claims. The 3,000-local fixtures improve from about 172 ms to 1.2–1.5 ms. Small scalar controls are mixed (100 locals: 0.293 to 0.304 ms; 3,000: 0.904 to 0.994 ms), and timing varies with process/machine load. The deterministic normalization budget guards against rebuilding host facts per statement without relying on timing assertions.

## Validation

- Three new tests cover bounded host normalization work, procedure shadowing with case-insensitive reads, and existing late-bound Collection/Word condition diagnostics.
- Eight focused files: 65 tests passed.
- Typecheck passed.
- Full suite: 604 files, 12,624 tests passed, 13 skipped.
- 1,500 complete analyzeModule outputs exactly match baseline 0928416d, including messages, spans and order. Five hosts (Excel, Word, PowerPoint, unknown and VB6), ten declared types, ten read/write forms, three module/local/shadowing configurations, alternating VBA7 activity; no internal errors in either build.
