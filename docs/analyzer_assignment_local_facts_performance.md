# Assignment direct local lookup

Issue: #826. Baseline: `946795a4c0f38e61d8db5497ebb1cf142b91edad`.

Three assignment value helpers independently searched the same procedure child declarations by lowercased name: arrayValueAt, nullHeldAt and knownStringAt. Repeated reads of a late local repeated those linear scans, sometimes in multiple helpers for the same statement. The helpers now share a lazy query cache per procedure within checkAssignmentTypes. Missing names are cached too. The first matching direct child remains authoritative; kind, Static, array and type guards still inspect that symbol as before. The broader binding resolver is not substituted for this direct-child predicate.

Only symbol lookup is cached. Held Null, array/string values, reaching assignments, mutation and statement spans remain statement-specific. The cache is discarded between procedures/invocations and is not keyed globally by parsed nodes.

Run `node scripts/benchmark-assignment-local-facts.mjs`, then with `--baseline=946795a4c0f38e61d8db5497ebb1cf142b91edad`. Three warmups, fifteen samples on Node 24.18.0 / Ryzen 7 9800X3D. The parser output is reused; each sample builds fresh symbol roots outside timing, with an identity assertion. Timing covers the assignment-type rule and its setup. No other audit tests/benchmarks ran during measurement. Each fixture has the indicated unrelated locals before v and n, followed by 1,000 reads. Fixed-string repeated reads retain the existing mention-count guard, so that case does not report a never-assigned fixed-string diagnostic. The separate single-read regression covers that diagnostic. These are isolated rule timings, not whole-analyzer speedups.

| Locals/case | Baseline median / p95 ms | Fixed median / p95 ms |
| --- | ---: | ---: |
| 10-variant-null | 3.171 / 4.406 | 2.924 / 3.568 |
| 10-variant-empty | 3.913 / 4.445 | 3.557 / 5.201 |
| 10-fixed-string | 3.09 / 3.774 | 2.854 / 3.971 |
| 10-known-string | 3.452 / 4.704 | 3.622 / 4.466 |
| 10-literal-control | 1.515 / 2.185 | 1.54 / 2.13 |
| 100-variant-null | 2.96 / 4.115 | 2.053 / 2.673 |
| 100-variant-empty | 4.186 / 5.404 | 2.865 / 3.746 |
| 100-fixed-string | 3.98 / 5.404 | 2.609 / 3.717 |
| 100-known-string | 3.647 / 4.28 | 2.717 / 3.592 |
| 100-literal-control | 1.622 / 2.313 | 1.527 / 2.429 |
| 1000-variant-null | 10.966 / 11.454 | 2.375 / 6.542 |
| 1000-variant-empty | 18.556 / 24.708 | 4.263 / 5.607 |
| 1000-fixed-string | 17.523 / 18.46 | 3.759 / 5.521 |
| 1000-known-string | 12.612 / 13.528 | 3.806 / 4.645 |
| 1000-literal-control | 2.009 / 3.882 | 2.062 / 3.544 |
| 3000-variant-null | 29.989 / 31.098 | 2.759 / 4.199 |
| 3000-variant-empty | 52.321 / 66.03 | 8.098 / 9.037 |
| 3000-fixed-string | 51.914 / 55.306 | 7.649 / 8.344 |
| 3000-known-string | 36.68 / 40.061 | 8.039 / 8.427 |
| 3000-literal-control | 3.868 / 6.435 | 3.826 / 5.434 |

Small controls are mixed: the 10-local known-string median rises 3.452 -> 3.622 ms, and the 1,000-local literal control rises 2.009 -> 2.062 ms. The improvement removes repeated declaration searches; it does not remove other inference/dataflow costs. Only queried names occupy the cache.

Actual symbol-name getter counts for 100 unrelated locals and 100 reads:

| Case | Before | After | Diagnostics |
| --- | ---: | ---: | ---: |
| Held Null | 20,923 | 722 | 100 |
| Empty Variant | 31,941 | 1,742 | 0 |
| Fixed String | 31,931 | 1,732 | 0 |
| Known String | 21,836 | 1,737 | 100 |

Validation: type checks; 246 focused tests; full suite (613 files, 12,724 tests passed, 13 skipped); 2,000 complete analyzer outputs matched across eight local declaration forms, parameters, module/local shadows, six target types, seven held values, eight read/statement forms, four hosts, conditional activity and multiple procedures. The subsequently added missing-local regression passed together with all ten tests in its file, followed by type checks. Tests assert actual read bounds, cached misses for module-only names, first-match/Static/parameter guards, procedure isolation, conditional refresh, statement-specific held-value changes and retained array/fixed-string diagnostics.
