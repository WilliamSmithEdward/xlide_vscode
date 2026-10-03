# Bulk canonical casing symbol reuse

Bulk canonical casing resolved every identifier independently. Identifier completion rebuilt the complete module symbol graph on each such query, even though all offsets shared the same source and module context.

A request-local identifier resolver now builds symbols lazily on the first applicable position and reuses them at subsequent offsets. It still selects the enclosing procedure for each offset and produces new completion arrays and records. Each bulk casing invocation gets a separate resolver; there is no new persistent cache. Single-position casing retains its existing behavior.

## Measurement

Run node scripts/benchmark-canonical-casing.mjs --rounds=15 in the target checkout. The runner bundles local source using the existing esbuild dependency, performs three warmups and reports 15-sample median/p95 times. Compare against commit 0cb744c2 using the same runner. Both runs were taken without the test suite running.

Node 24.18.0, AMD Ryzen 7 9800X3D, Windows. Bulk fixtures contain repeated lowercase references to a local declared as Value and require two casing edits per statement. Host/runtime suggestions are disabled to isolate source-symbol work.

| Fixture | Before median ms | After median ms | Before p95 ms | After p95 ms |
| --- | ---: | ---: | ---: | ---: |
| 10-procedures/10-statements | 3.3 | 1.965 | 4.893 | 3.277 |
| 100-procedures/10-statements | 227.502 | 115.921 | 243.331 | 147.012 |
| 1-procedures/1000-statements | 136.641 | 69.729 | 142.151 | 79.298 |
| single-position/100-procedures | 0.144 | 0.091 | 0.253 | 0.295 |

The large bulk fixtures are approximately 49% faster. These are warm synthetic measurements, not a general workbook latency guarantee. The single-position control remains sub-millisecond; no speedup is claimed for that control.

## Validation

The symbol-build regression reproduces 60 builds for 30 statements against the original bulk casing implementation and requires one build after the change. Additional coverage checks per-procedure local scopes, independently owned completion records, lazy declaration-name refusal, and isolation across source revisions. Existing identifier/casing/controller tests cover their completion context and generated edits.
Type checking passed. The final full Vitest suite passed: 557 files, 11,983 tests passed and 13 skipped.
