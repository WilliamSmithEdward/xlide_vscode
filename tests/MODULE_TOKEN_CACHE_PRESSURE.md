# Module lexer cache pressure

Short expression and declaration lookups shared eight lexer-cache slots with
whole modules. A burst of short lookups could evict the active large class,
leaving completion and semantic tokens to lex the entire class on the next edit.

The lexer now keeps separate least-recently-used caches for sources shorter
than 4,096 characters and sources at or above that size. Each cache retains
at most eight entries. The threshold matches the existing incremental-lexing
threshold. Exact source equality, token coordinates, canonical spelling,
trivia, read-only snapshots and full-lexer fallback behavior are preserved.
There are at most sixteen cache records overall; small string slices can
retain a larger parent string, so this is an entry-count bound, not a fixed
heap-size guarantee.

The public regression covers complete semantic results on six source revisions
under twenty short lookups per edit, LF/CRLF/CR metadata parity, deterministic
mixed edit histories, threshold crossings, newline fallback, equal-length
sources with identical ends, copied equal strings, immutable old snapshots,
retention in both directions, and eventual eviction from each cache.

A real 64-cycle workbook run originally logged 65 full-class lexer misses in
the extension host: one cold parse, 39 completion requests and 25 semantic
requests. Every miss had no compatible prior module; the cache held only short
strings. With the split cache, the matching diagnostic run completed all twelve
cases and recorded 37 host misses. Only the initial cold parse lacked a prior
module. The other 36 misses had a compatible snapshot and a changed span that
crossed an existing newline, taking the conservative full-lexer fallback.
Instrumentation recorded lengths and static stack frames without module text.
These diagnostic runs establish the eviction cause; logging timings are not
used as controlled UI-speed measurements.

The renderer harness now reveals the variable's statement before mouse hover,
chooses a visible occurrence rather than an offscreen buffered Monaco row, and
restores and verifies the member caret afterward. It still requires a resolved
`LatencyValue As Long` tooltip and subsequent keyboard-driven completion.

The final analyzer comparison used identical dependencies and semantic code,
with the old versus split-cache lexer bundled separately. Each trial interleaved
50 warm requests and 60 fresh fixed-width comment revisions per variant. Module
index updates and ten short pressure lookups occurred outside the query timer.
All four semantic collectors were timed together, and complete results matched
exactly across variants and revisions. The private class workload was 969,469
characters; its source and profiles remain in ignored local artifacts.

Times below are median / p95 / maximum milliseconds. Warm and unpressured
request medians remain similar; individual warm and unpressured tails do not
all improve. The reliable work reduction is under short-lookup pressure.

| Complete semantic request | Trial | Original cache | Split cache |
| --- | --- | --- | --- |
| Warm, no pressure | 1 | 9.04 / 11.48 / 16.64 | 8.43 / 10.72 / 12.19 |
| Warm, no pressure | 2 | 8.87 / 12.19 / 14.54 | 8.65 / 15.15 / 29.33 |
| Fresh revision, no pressure | 1 | 15.37 / 22.34 / 23.96 | 15.72 / 24.70 / 33.20 |
| Fresh revision, no pressure | 2 | 15.26 / 20.04 / 36.31 | 14.82 / 20.01 / 21.37 |
| Fresh revision, short pressure | 1 | 40.06 / 58.19 / 85.89 | 15.12 / 21.91 / 25.91 |
| Fresh revision, short pressure | 2 | 39.62 / 47.61 / 68.22 | 15.16 / 25.25 / 28.86 |

Each pressured revision lexed 969,469 characters in the original cache and a
50-character edit window with the split cache. Over 60 revisions in each
trial, that was 58,168,140 versus 3,000 freshly lexed characters. Indexing is
excluded from these query timings, and they do not measure menu or hover paint.

Production validation passed type checking and compilation, 15,316 unit tests
(33 existing skips), and 39 real editor cases covering completion acceptance,
miss recovery, bracketed members, hover, semantic tokens, references, casing,
signature help, Enter and document lifecycles. The final unprofiled private
workbook run passed all twelve integration cases with 64 fresh-source renderer
cycles and four resolved mouse hovers.

| Unprofiled editor observation | Median | p95 | Maximum |
| --- | --- | --- | --- |
| Backspace paint | 16 ms | 31 ms | 32 ms |
| Completion menu paint after deletion | 125 ms | 145 ms | 153 ms |
| Typing paint | 31 ms | 31 ms | 33 ms |
| Miss menu clearing | 1 ms | 1 ms | 11 ms |
| Resolved mouse hover | 379 ms | 395 ms | 395 ms |

The final profiled run passed all twelve integration cases, 1,000 fresh-source
renderer recovery cycles and 62 resolved mouse hovers. Both process profiles,
timings, observations and the exact compiled extension bundle/source map were
retained with verified SHA-256 digests. No temporary lexer logging was present
in the final production build.

| Profiled editor observation | Median | p95 | Maximum |
| --- | --- | --- | --- |
| Backspace paint | 16 ms | 46 ms | 115 ms |
| Completion menu paint after deletion | 125 ms | 154 ms | 207 ms |
| Typing paint | 16 ms | 40 ms | 73 ms |
| Miss menu clearing | 1 ms | 23 ms | 112 ms |
| Resolved mouse hover | 369 ms | 388 ms | 416 ms |

Slow outliers remain. The 115 ms Backspace window overlapped approximately
38 ms of renderer garbage collection and 25.1 ms of local project-context
work in the host. The 112 ms miss-clearing window was mostly idle in both
sampled processes. These approximate profile overlaps do not establish a
single cause. Sampled inclusive host work still included 12.92 seconds in
`tokenizeCached`, 12.21 seconds in `macroNameStringAt`, 14.85 seconds in semantic
tokens and 18.57 seconds in canonical casing over a 504.86-second capture.

The ordinary and profiled observations demonstrate working keyboard recovery
and resolved hovers on this workload. The controlled analyzer comparison
establishes the cache-pressure improvement. Intervening main changes,
profiling, renderer scheduling and the repaired hover navigation limit any
comparison of overall UI speed with earlier captures. Multi-line lexical
fallback and repeated project-context construction remain follow-up targets.

The public cache regressions can be run with
`npx vitest run tests/moduleTokenCachePressure.test.ts`. The full unit suite uses
`npx vitest run --maxWorkers 2`. For the optional workbook/renderer harness,
compile first, provide a local large-class workbook in `XLIDE_PERF_WORKBOOK`,
set `XLIDE_PERF_FRESH_SOURCES=1` and `XLIDE_PERF_RENDERER_CYCLES=1000`, and use
`npx vscode-test --config tests/nativeBackspaceIntegration.config.mjs` with the
pinned Code installation. `XLIDE_PERF_CPU_PROFILE=1` retains both process
captures. The harness edits a disposable workbook copy.
