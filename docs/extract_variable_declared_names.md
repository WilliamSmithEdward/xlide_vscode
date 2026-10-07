# Extract Variable declared names

Extract Variable's generated-name allocator scanned the procedure text with a declaration regex. It reserved only the first bare name after Dim/Static/Const/ReDim, missing later grouped names, bracketed declarations and Preserve array targets. It also reserved fake declarations inside comments and strings. A procedure containing `Dim first As Long, value As Long` therefore received a duplicate `Dim value As Double` when extracting `2 * 3`.

Read every name from the existing immutable parsed variable groups, including nested bodies. For raw statements and their single-line If branch spans, lex only anchored declaration/ReDim candidates using shared cached statement tokens, leading-line-number removal and top-level comma grouping. Preserve implicit ReDim names, including Preserve and multidimensional bounds, while leaving qualified object/member targets out of the local-name set. Remove the full-procedure declaration regex. The final free-name search and module/parameter reservation remain unchanged. No persistent cache is added.

Thirty-six independent complete-result cases cover grouped Dim/Static/Const declarations, bracketed/case-insensitive names, continuations, nested blocks, implicit ReDim/Preserve groups, numbered and single-line conditional statements, comments/strings, qualified targets and free-name controls across LF/CRLF. Twenty-eight fail on baseline and eight already pass; all 36 pass after restoring exact production bytes. Types, 62 focused checks and the full suite pass: 14,848 tests, 33 skipped, 763 passing files, seven skipped files, 73.04 seconds.

All 16,512 complete results match baseline using 8,256 oracle sources and two newline styles; 16,496 queries succeed. Derive the first parsed procedure's body, or module variable groups when there is no procedure, and insert an independent numeric extraction. This is compatibility coverage for those derived bodies, not arbitrary malformed-module compilation or a corpus containing the newly corrected collision shapes. The new tests provide those independent expectations.

## Timing and costs

Run `node scripts/benchmark-extract-declared-names.mjs` in baseline/current/current/baseline order, with `--baseline=878faab4b92f9e45b0f022cfb59dff04ea5a50fa` on baseline runs. Node 24.18.0, Ryzen 7 9800X3D; three warmups, nine measured rounds, ten complete refactors per round. Source/AST warming is outside the clock. Complete applied output, title and rename selection are independently checked afterward.

This is a correctness repair with mixed performance effects. Parsed single declarations and trivia-heavy cases improve; traversing ordinary statements and tokenizing ReDim groups costs more. In 1,000-statement equal-output controls, bare bodies rise from 0.02076–0.02224 ms to 0.05346–0.07788 ms, and ReDim bodies from 0.07767–0.07825 ms to 0.17931–0.20459 ms. Tiny ReDim inputs also cost more. Grouped/trivia baseline fixtures intentionally assert the old incorrect names, so those are not equal-output timing comparisons. No cold-parser, heap-byte or editor-latency gain is claimed.

Ranges are medians from two independent runs, milliseconds per complete warmed query.

| Body statements | Kind | Baseline ms | Repair ms |
| --- | --- | --- | --- |
| 1 | bare | 0.01243–0.01313 | 0.01283–0.01319 |
| 1 | single | 0.00868–0.00872 | 0.00786–0.00817 |
| 1 | grouped | 0.00524–0.00539 | 0.00568–0.00618 |
| 1 | trivia | 0.00577–0.00730 | 0.00554–0.00588 |
| 1 | redim | 0.00562–0.00585 | 0.00751–0.00995 |
| 1000 | bare | 0.02076–0.02224 | 0.05346–0.07788 |
| 1000 | single | 0.11901–0.12016 | 0.08384–0.08628 |
| 1000 | grouped | 0.12396–0.16549 | 0.10525–0.10645 |
| 1000 | trivia | 0.07438–0.07575 | 0.00899–0.01042 |
| 1000 | redim | 0.07767–0.07825 | 0.17931–0.20459 |
