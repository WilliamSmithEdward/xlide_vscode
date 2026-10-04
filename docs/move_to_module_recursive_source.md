# Move to Module: moved references and bounded masking

Move to Module repointed qualified recursive references in the old module's removal range, while copying the unchanged procedure to its destination. The inserted procedure therefore still called Reports.Build after Build had moved to Helpers. The primary edit list also contained overlapping removal/replacement spans. Different-length destination names could corrupt the original source when applying those overlapping edits.

The procedure's qualified references are now repointed in the copied text before insertion. The primary source receives only references outside its removal span. Attached documentation, strings, comments, argument text and retained procedures keep their original source fidelity. Target-existing caller edits and the new procedure insertion remain one atomic module edit set.

The qualified-reference scanner previously located and masked the entire physical line for every match. One colon-separated line with 1,000 calls caused 14,998,000 characters to pass through stripVba. Ascending matches now reuse the same query-owned physical-line bounds and mask: 14,998 characters for that line. The obsolete per-occurrence helper is removed. Unmatched lines are not masked, and no lexer/module cache is populated by this mask.

## Validation

Eighteen new tests cover recursive bare/Call forms, multiple colon calls, spacing/case, different-length destination names, attached docs, comments/strings and primary/target full text across LF, CRLF and CR. Six work cases span 1, 100 and 1,000 calls on one or many physical lines. Baseline 82de7a08 fails 14 and passes four controls. Candidate focused validation passes 77 tests across five files; types pass.

With frozen parser ASTs and lexer tokens/trivia, 8,256 corpus sources produce 16,500 complete Move results and full primary/destination outputs identical to baseline. Those corpus queries do not exercise the specific qualified recursion. An additional 72 independently expected full refactors vary recursive form/count, EOL and empty/occupied target modules: baseline matches 24, candidate matches all 72. Every candidate edit set has explicit bounds and non-overlap checks; there are no uncaught exceptions.

## Measurement

Run node scripts/benchmark-move-to-module.mjs --rounds=9, with --baseline=82de7a08 for the old implementation. Node v24.18.0, AMD Ryzen 7 9800X3D 8-Core Processor           ; three warmups, nine timed rounds, separate sequential processes in baseline/candidate/candidate/baseline order. Cached source is unchanged; fresh source varies a leading comment. Timing covers the pure Move operation; edit application and complete primary/target/caller verification are outside the timed interval. All 18 layouts produce correct complete outputs before and after. Ranges span two process medians; no Office/editor latency is measured.

| Calls / layout / source | Before median range (ms) | After median range (ms) |
|---|---:|---:|
| 1/colon/cached | 0.018–0.018 | 0.019–0.021 |
| 1/colon/fresh | 0.005–0.005 | 0.005–0.005 |
| 1/multiline/cached | 0.005–0.005 | 0.005–0.005 |
| 1/multiline/fresh | 0.005–0.005 | 0.005–0.005 |
| 1/comments/cached | 0.008–0.008 | 0.006–0.006 |
| 1/comments/fresh | 0.006–0.011 | 0.006–0.008 |
| 100/colon/cached | 1.287–1.323 | 0.035–0.038 |
| 100/colon/fresh | 1.361–1.374 | 0.020–0.021 |
| 100/multiline/cached | 0.025–0.044 | 0.044–0.050 |
| 100/multiline/fresh | 0.025–0.046 | 0.024–0.028 |
| 100/comments/cached | 0.067–0.067 | 0.037–0.041 |
| 100/comments/fresh | 0.070–0.074 | 0.040–0.043 |
| 1000/colon/cached | 123.094–123.133 | 0.161–0.164 |
| 1000/colon/fresh | 123.225–123.910 | 0.172–0.183 |
| 1000/multiline/cached | 0.201–0.201 | 0.191–0.205 |
| 1000/multiline/fresh | 0.202–0.203 | 0.182–0.184 |
| 1000/comments/cached | 0.593–0.622 | 0.312–0.317 |
| 1000/comments/fresh | 0.620–0.638 | 0.331–0.332 |


This patch retains the existing qualified-name matching/binding behavior. Bracketed and continued receiver forms, shadowing, visibility and destination-name capture still require the remaining Move surface audit. Full repository validation is running before marking this fix ready.
