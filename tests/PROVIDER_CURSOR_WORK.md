# Completion provider cursor classification

## Finding and change

The provider and Backspace recovery used the full-prefix cursor API to inspect
comments, strings, a bracketed member's preceding token and the current
statement boundary. These checks need only the current logical line. Each new
source/caret pair previously copied and filtered every earlier token, even
when a comment or ordinary string immediately rejected completion.

All seven provider call sites now use the existing logical-line cursor API.
Project lookup, completion contents, incomplete-list refresh and keyboard
recovery behavior are unchanged. Continued comments, strings, bracketed names
and statement boundaries retain their existing classification.

## Regression evidence

Four work-count cases near the end of a 1,200-procedure module failed against
the previous provider with 15,622-15,641 cached-token reads. They pass with fewer
than 100 reads after the change. They cover ordinary comments, ordinary strings,
a closed macro-name string and recovery inside a continued comment. A separate
case covers a comment continued onto the next physical line. Existing all-offset
prefix-window equivalence and completion/recovery suites pass.

Compilation and the full local suite passed: 14,673 tests across 750 files,
with 33 skipped tests. The suite used four workers. After rebasing onto main's host member-name index
reuse (#1100), compilation, 73 targeted tests, 25 completion editor cases and
all 12 native/workbook cases with 64 fresh-source cycles passed again.

## Private component comparison

On a 969,415-character fixture derived from the private class, the two isolated
bundles use the same cursor implementation and differ only in which cursor API
is called. Sixty warm checks at ten varying offsets measured median/p95/max
2.04/5.87/7.08 ms for full-prefix classification and 0.0087/0.0132/0.0211 ms for
logical-line classification. Forty fresh-source checks, changing a fixed-width
synthetic comment, measured 4.19/7.10/7.54 ms before and 2.04/5.73/11.44 ms after.
Fresh-source lexing and occasional slow samples remain; these component results
do not establish menu painting latency or eliminate every source of typing lag.
Private source, bundles and profiles stay in ignored .vscode-test.

## Renderer validation

All 12 native/workbook cases and 25 completion editor cases passed. The
profiled repeat passed 800 fresh-source cycles and 50 mouse hovers on pinned
Code 1.139.1. Before/after menu median/p95/max was 117/136/193 ms and
111/133/152 ms. Early 32-cycle menu median/max was 132/193 ms and 110/135 ms;
later-cycle median/max was 116/166 ms and 113/152 ms. The comparison also
includes main's unrelated intervening fixes (#1096 and #1098), so it does not
isolate the provider change's effect on UI timings.

After-change visible Backspace median/p95/max was 16/32/102 ms, typing
16/44/48 ms and miss invalidation 1/12/26 ms. Mouse hover was 377/400/410 ms,
including Code's hover delay. Host heartbeat maximum was 77.21 ms, compared
with 114.56 ms before. There was no lost deletion or unresolved hover. The
102 ms Backspace sample overlapped renderer work and remains a slow sample;
these measurements include debugger/polling overhead, rather than GPU paint.
An unprofiled repeat at e2979936 passed all 12 native/workbook cases, 1,000
fresh-source cycles and 62 mouse hovers. Visible Backspace median/p95/max was
16/35/69 ms, typing 18/47/61 ms, menus 119/139/328 ms, miss invalidation
1/12/82 ms and hover 382/413/419 ms. Command-driven typing median/max was
1.73/10.12 ms and Backspace 1.73/18.02 ms, with a 23.26 ms host-heartbeat
maximum in the separate sustained command test. The longest menus were cycles
985 (328 ms) and 997 (307 ms), both after 30 ms idle. The renderer stress test's
host heartbeat is currently recorded only while profiling; these unprofiled
popup outliers therefore lack a correlated host trace.

The earlier missing-Backspace observation at cycle 113 and the unprofiled
651 ms menu outlier did not recur in either repeat. They remain
unclassified; this component fix does not close the overall stalling goal.
