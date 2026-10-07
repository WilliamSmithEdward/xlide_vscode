# Procedure label fact reuse

Declaration and reference collectors walked and token-sliced the same procedure for every reader. Ten paired readers of a 1,000-statement procedure performed 20,040 statementTokensCached requests, even though source, procedure and conditional activity were unchanged.

A weak procedure cache now retains lazy declaration/reference arrays for its latest source/activity pair. Asking for one kind does not compute the other. Changed source or tracker identity replaces both fields; equal-source hits adopt the current caller string. Every returned array contains cloned labels and spans, preserving caller mutation isolation. Direct statement helpers and module-wide target semantics are unchanged.

Retained memory is proportional to current labels/references per live procedure. Empty collections are retained too, so repeated no-label queries avoid walking all statements. This does not add per-source version chains or memoize arbitrary callbacks; activity trackers are immutable analysis-context objects as in the existing flow cache.

Validation against ef8a34257737009e6d9ed1fbf9e3f3bf534be856:

- Types and 117 focused tests across five files pass.
- Three baseline work-count failures: 240/2,040/20,040 scanner requests for 10/100/1,000 repeated statements with ten paired readers. After the fix, no more than two initial scans (24/204/2,004 calls) are needed. Output mutation isolation and source/activity replacement controls pass on both versions. The initial fixture's label spelling expectation was corrected before the genuine work-count baseline run.
- Frozen-input differential over 8,256 corpus sources plus 36 generated cases: complete declaration/reference arrays (including keys, spellings, spans and reference kinds), equal-source warm results and all 8,292 complete diagnostic arrays match (13,440 findings), with zero internal errors. Generated cases cover LF/CRLF/CR, true/false activity, named/numeric/combined labels, handler targets, computed GoTo and Resume. ASTs, tokens and trivia are frozen.

Full suite: 731 files passed, seven skipped; 14,483 tests passed, 32 skipped. No failures.

Benchmark: node scripts/benchmark-procedure-label-facts.mjs --baseline=ef8a34257737009e6d9ed1fbf9e3f3bf534be856 --rounds=9 and without baseline. Baseline/candidate/candidate/baseline order, three warmups/nine measured rounds; Node 24.18.0 on Ryzen 7 9800X3D. Every round uses unique source, with one named label and GoTo around the repeated statements. Both initial collections are warmed. Timing covers ten paired declaration/reference readers, excluding parsing/construction. Every complete object, including independently calculated absolute spans, is checked.

| Repeated statements | Before median ms / ten readers | After median ms / ten readers |
| --- | --- | --- |
| 10 | 0.0316–0.0364 | 0.0044–0.0045 |
| 100 | 0.1802–0.1858 | 0.0046–0.0050 |
| 1,000 | 1.0058–1.0510 | 0.0037–0.0042 |

These component measurements do not establish total editor latency. Initial collections still walk the procedure, returned objects are still copied, and label-dense procedures incur output-copy cost proportional to their labels/references.
