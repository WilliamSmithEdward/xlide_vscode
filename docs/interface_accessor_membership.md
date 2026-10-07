# Interface accessor membership and bounded matching

Implement Interface indexed existing procedures by name alone. A Property Get therefore satisfied a required Get/Let or Get/Set pair, and an incompatible Sub/Function could be reported as a completed member. This leaves required interface members missing.

The refactor now matches case-insensitive qualified names and exact procedure kinds. Get/Let/Set can share a name and are checked independently. A different Sub/Function/property family is refused before adding a conflicting declaration. Existing signatures are copied as before, and fully implemented classes retain their refusal message. This does not validate the signatures or visibility of every pre-existing implementation.

The query-owned index retains only procedure names required by the selected interface, instead of indexing every unrelated class procedure. It still scans class members once. Existing project source gathering and cold parsing are outside this change; there is no new cache or response wait.

Thirty-seven new tests cover scalar and object fields, explicit accessor pairs, either missing leg, three newline styles, case-insensitive interface/member names, complete pairs, incompatible callable kinds, wrong writer kinds and unrelated-name indexing. Twenty-nine semantic cases fail on the initial baseline. All 71 focused checks and types pass. The validated header repair from #1120 is included when running the combined tests so CR signatures are preserved.

## Measurement

Run `node scripts/benchmark-interface-accessor-membership.mjs` in baseline/current/current/baseline order, supplying `--baseline=28c54edb92a928b05130ae5212d4f04a99ee0b65` for baseline. Node 24.18.0, AMD Ryzen 7 9800X3D; three warmups and nine measured rounds. Twenty complete refactor calls per round; class/interface ASTs and source construction outside the clock, exact result checks afterward.

The partial baseline result is the known incorrect fully-implemented refusal; the repaired result adds the missing setter. Generation with no existing accessors and the fully complete controls have identical full results. No cold-parser, heap-byte, command-latency, renderer-paint or universal speedup claim. Matching still scans the class once and retains kinds only for required names. Tiny cases can be mixed.

10,000 unrelated procedures; milliseconds per complete result. Ranges are medians of two independent runs.

| Scenario | Baseline | Repair |
| --- | --- | --- |
| partial | 0.66881–0.71352 | 0.30168–0.30494 |
| none | 0.66895–0.68274 | 0.28227–0.28952 |
| complete | 0.66568–0.70446 | 0.27925–0.29644 |

## Complete validation

All 37 new tests pass. Baseline rerun fails 31 cases and passes six complete-pair controls; the two work-count cases prove unrelated names are not indexed. Types, 71 focused tests and 14,796 full-suite tests pass (33 skipped, 761 passing files, seven skipped files, 75.24 s). Independent expectations match 576 complete scalar/object property results across field/explicit declarations, missing-leg patterns and three newline styles. All 16,512 compatible complete corpus results match baseline. These comparisons prove tested edits/refusals, not actual Office compilation or correctness of every pre-existing implementation signature.
