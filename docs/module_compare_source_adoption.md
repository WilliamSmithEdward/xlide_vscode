# Module Compare caller-source adoption

moduleCompare retains the most recent source and its Binary/Text/Database result. On an equal-content cache hit it previously retained the older source string. If an editor caller supplies a separate equal-content string and then reuses that string for many procedures, each call can compare the entire source against the older string. Retaining the current caller string after equality is established makes subsequent calls use the same string representation.

The cached mode and miss path are unchanged. The cache still owns one source/result pair; replacement releases the older string rather than retaining a history. Same-string calls already avoid content comparison and are the benchmark control. No total editor latency or cold-scan improvement is claimed.

Validation against 49ed5c90fc93893cc8f26beb0d9e5c299e3523e4:

- Types and 37 focused tests across known string calls, assignments and Split/Filter comparisons pass.
- Frozen-input differential: 8,256 corpus sources plus 36 generated cases (LF/CRLF/CR, Binary/Text/Database/default and 1/10/100 comparisons). All 8,292 complete diagnostic arrays match (14,095 findings), with zero internal errors. Each version also independently matches the diagnostics for original and equal-copy source strings, and all repeated mode answers agree. ASTs and lexer tokens/trivia are frozen.

Full suite: 728 files passed, seven skipped; 14,455 tests passed, 32 skipped. No failures.

Reproduce with node scripts/benchmark-module-compare.mjs --baseline=49ed5c90fc93893cc8f26beb0d9e5c299e3523e4 --rounds=9 and without baseline. Baseline/candidate/candidate/baseline order, three warmups/nine rounds, Node 24.18.0 on Ryzen 7 9800X3D. Every mode answer is independently checked. Construction and priming are excluded; a unique comment for each round forces priming to miss, so the same-string control really starts with the current string. Approximate sizes include filler plus an option/comment header. Each measurement includes 1,000 cache hits and result assertions.

| Approximate source size | Equal-copy before ms | Equal-copy after ms | Same-string before ms | Same-string after ms |
| --- | --- | --- | --- | --- |
| 1 KB | 0.0583–0.0585 | 0.0275–0.0291 | 0.0429–0.0485 | 0.0396–0.0429 |
| 100 KB | 1.5756–1.6791 | 0.0139–0.0141 | 0.0167–0.0173 | 0.0124–0.0125 |
| 1 MB | 16.7031–16.8473 | 0.0248–0.0295 | 0.0049–0.0050 | 0.0123–0.0124 |

Table values are Text-mode medians; Binary and Database answers also pass all 18 benchmark configurations. The same-string 1 MB control adds about seven nanoseconds per hit in this run, consistent with the added assignment. Equal copies require one content comparison before adoption; alternating different strings still miss. This synthetic component workload does not measure how often real editor callers supply equal copies or their total latency.
