# Active-call cursor context

An extension-host profile from the private large class still showed repeated
`findParenCall` work during typing. Active-call detection and callable completion
parenthesis decisions requested the complete token prefix and then walked it,
though newline/colon boundaries already discard earlier statements.

Both helpers now request the existing logical-line cursor window. It preserves
absorbed line continuations, truncated strings/comments/operators, numeric
labels, absolute callee offsets and the same classifiers. Previous statements
no longer create temporary parenthesis frames or enter each call-site scan.

Regressions compare the complete old full-prefix results at every caret through
nested, continued, member/parenless, colon-separated, labelled, Unicode, `$`
function, string/comment, directive and malformed-call samples. A token-visit
probe checks fewer than 80 reads after 24,000 unrelated call lines, and rejects
any use of the full-prefix API. The optional private-corpus check compares 66
positions without logging source (`XLIDE_INCREMENTAL_PARSE_CORPUS`).

On the 969,328-character private snapshot, alternating 21 batches of 100 warm
active-call checks measured a median batch mean of 2.226214 ms with the old
prefix stream and 0.000419 ms with the logical-line stream. Maximum batch means
were 2.561875 and 0.004108 ms respectively. This is a warmed component comparison,
not an overall editor speedup claim; lexing and other analyzer work still have
separate costs.

The first actual-workbook repeat passed all eight cases, including the renderer
and busy-host controls. The 48 command typing/Backspace pairs had medians
2.38/2.03 ms and maxima 25.11/23.23 ms, with a 46.42 ms maximum host heartbeat
delay. Visible typing and Backspace maxima were 46/32 ms across 24 recovery
cycles, each displaying Cells and then clearing it on the subsequent `.cez`
miss. The first menu show took 117 ms after deletion; the installed VS Code
1.139.1 workbench schedules its first `visible` class update with a 100 ms timer,
so this includes editor widget scheduling as well as provider delivery. Updated
declaration hover returned in 29.93 ms. Startup, profiling and run scheduling
are reported separately from warm component timing.
