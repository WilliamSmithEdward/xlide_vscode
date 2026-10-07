# Extract Method procedure-scoped identifier scan

Tested parent: `203e2eca06f9504cdd00c8689d4c7b8fb86ecb20` (PR #1229).

Local classification previously collected matching identifiers throughout the
module, then removed occurrences outside the containing procedure. Repeated
extraction in a small procedure therefore swept every unrelated procedure,
allocated matching occurrences, and discarded them. The existing scanner
already supports inclusive absolute offset bounds.

Pass the containing procedure's span to that scanner and remove the redundant
post-scan range filter. This retains the same scope and absolute occurrence
coordinates. Whole-source stripping still supplies continuation/comment context;
only the identifier sweep is bounded. Raw first-occurrence parameter ordering
still uses the complete source, including comments and shared substrings.

## Measured work and timings

Run `node scripts/benchmark-extract-procedure-scan.mjs` and compare with
`node scripts/benchmark-extract-procedure-scan.mjs --baseline=203e2eca06f9504cdd00c8689d4c7b8fb86ecb20`.
The benchmark includes 0, 100, 1,000 and 10,000 unrelated assignment statements
before or after a selected Sub/Function. Every physical line is below 1,023
characters. A separate untimed bundle counts identifier-scan physical lines;
untouched bundles measure the public API, with and without applying edits.

Four independent processes ran before/after/after/before. Each case had three
warmups and nine samples; medians are compared below. Complete output digests
matched in every process, with independent helper/invocation assertions. Runtime:
Node 24.18.0, AMD Ryzen 7 9800X3D.

| 10,000-statement case | Before API median ms | After API median ms | Identifier lines before → after |
| --- | --- | --- | --- |
| Sub, unrelated procedures before | 0.6858–1.2636 | 0.0085–0.0129 | 10,067 → 5 |
| Sub, unrelated procedures after | 0.6723–0.7354 | 0.0101–0.0148 | 10,067 → 5 |
| Function, unrelated procedures before | 0.7411–0.7586 | 0.0124–0.0205 | 10,069 → 7 |
| Function, unrelated procedures after | 0.6805–0.6851 | 0.0091–0.0095 | 10,069 → 7 |

Including edit application, these large cases measured 0.6693–1.2979 ms before
and 0.0111–0.0395 ms afterward. Small controls have mixed/overlapping timings;
this is not a universal speedup claim. Repeated-source cache behavior is included.
No native/editor latency, cold-start, heap, or complete repository claim is made.

## Validation

- Type checking and 15 dedicated tests pass under LF, CRLF and CR. Three work
  invariants fail on the parent (4,010 unrelated regex sweeps each); 12 output
  controls pass on both versions.
- 29,268 complete public-result and applied-text comparisons have zero differences:
  500 generated modules, 1,000 Function controls, 3,000 mixed local/procedure
  contexts, and 24,768 corpus queries across 8,256 environments with three line
  endings. Mixed contexts include primitive, Variant, object and array locals,
  Unicode names, and matching local names in procedures before/after the caller.
- All 16,274 tests across 815 files pass (33 tests / seven files skipped).
  The invocation allowed 15 seconds per test, matching the parent validation.
- Generated sources from all 30 previously validated original/generated native
  pairs (60 executions) remain byte-identical. No new native executions were
  needed for this source-equivalent scan optimization.
