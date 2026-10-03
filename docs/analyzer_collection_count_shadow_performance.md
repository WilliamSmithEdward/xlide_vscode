# Collection.Count project-name predicate

A Collection.Count assignment guard rejects writes to the intrinsic Count Function unless a project surface named Collection exists. Previously each eligible assignment scanned projectClassMembers with the same case-insensitive name predicate. The rule now resolves that predicate lazily once per checkAssignmentTypes invocation and shares it across procedure member checks. Both true and false results are retained.

The predicate still matches any surface kind, with exact lowercasing and no qualification stripping or whitespace normalization. It preserves the original first matching scan, receiver-type/argument guards and all later completion behavior. Ordinary assignments do not request it. The result is invocation-local and refreshed for subsequent calls using the same mutable project metadata; no persistent cache, eager index or class-only completion lookup is added. Metadata is stable during an invocation.

## Deterministic cost

Actual project name getters measured all name reads in the direct rule, including completion work, with 100 surfaces and 100 assignments:

| Query fixture | Before name reads | After name reads | Diagnostics |
| --- | ---: | ---: | ---: |
| No Collection surface | 10,000 | 100 | 100 |
| Collection first | 202 | 103 | 0 |
| Collection last | 10,202 | 203 | 0 |
| Unused predicate | 0 | 0 | 0 |

The first/last fixtures use a writable project Count property. Their remaining name reads belong to other completion work; the cache does not remove that work. Instrumentation runs outside timing and adds no production counter.

## Reproduction and timings

From the repository root with installed dependencies:

```powershell
node scripts/benchmark-collection-count-shadow.mjs --baseline=30222875
node scripts/benchmark-collection-count-shadow.mjs
```

The baseline substitutes only assignments.ts from that commit in the current harness. Fixtures contain the listed number of class surfaces, plus a Collection surface at the start/end when applicable. Each procedure has 1,000 c.Count = 1 assignments with c declared As Collection, or scalar n = 1 assignments for the unused control. The benchmark checks 1,000 diagnostics only for the missing-shadow fixture and zero for matching/unused controls.

Parsing, tokenization and binding are outside timing. The parsed module and significant tokens are reused, while each sample receives a fresh bound root and fresh completion caches matching the diagnostic engine's pass lifetime. Bound-root identity is asserted. Measurements cover checkAssignmentTypes, not complete analyzer/editor latency. Three warmups precede 15 samples. Before and after ran sequentially after all audit tests and differential probes finished.

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor; measured locally on 2026-10-03. Times in milliseconds.

| Surfaces / mode | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: |
| 10-missing | 0.690 | 0.594 | 2.094 | 1.802 |
| 10-first | 2.649 | 2.649 | 3.657 | 3.066 |
| 10-last | 1.963 | 1.763 | 2.544 | 2.560 |
| 10-unused | 1.703 | 1.677 | 2.381 | 2.380 |
| 100-missing | 0.779 | 0.398 | 0.936 | 0.445 |
| 100-first | 1.790 | 1.692 | 2.124 | 1.973 |
| 100-last | 2.128 | 1.711 | 2.497 | 2.105 |
| 100-unused | 1.715 | 1.717 | 1.879 | 2.038 |
| 1000-missing | 4.245 | 0.317 | 4.572 | 0.481 |
| 1000-first | 1.733 | 1.691 | 2.111 | 2.019 |
| 1000-last | 6.202 | 1.684 | 9.172 | 2.949 |
| 1000-unused | 1.653 | 1.654 | 2.272 | 2.474 |
| 3000-missing | 12.929 | 0.350 | 14.618 | 0.613 |
| 3000-first | 1.768 | 1.684 | 1.977 | 2.129 |
| 3000-last | 14.343 | 1.682 | 14.971 | 2.327 |
| 3000-unused | 1.647 | 1.650 | 1.907 | 1.999 |

The 3,000-surface missing-shadow fixture improves about 37 times, and the last-match fixture about 8.5 times. First-match and unused medians remain close to baseline. Some small/control p95 results regress; these measurements do not establish a universal latency improvement. The first query still scans until a match or the end, and other member completion costs remain.

## Validation

- Type check passed.
- Focused assignment tests: 177 passed.
- Full suite: 618 files, 12,766 tests passed, 13 skipped.
- Twelve regressions cover actual read bounds across procedures, every surface kind, exact case/qualification/whitespace matching, unused metadata, mutable metadata refresh and receiver/argument guards.
- 1,500 differential cases preserve complete analyzer and direct assignment outputs across hit/miss metadata, duplicate surface kinds, accessors, Object/Long/VBA.Collection receivers, conditional declarations, nested With/branches/loops and multiple procedures.

Counting and generated differential probes remain audit scratch material; the timing benchmark and regression tests are committed.
