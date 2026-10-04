# Identifier completion and casing prefix work

## Finding

A CPU profile of 256 renderer-driven fresh-source edit cycles in the owner's
969,296-character class found about ten seconds in automatic casing over a
99-second run. Identifier lookup contributed about 5.4 seconds inclusively.
Its grammar checks copied and filtered every token before each identifier,
even though declaration/type/member gates, boolean literals and explicit Call
only inspect the current logical statement.

## Change and regression coverage

Identifier lookup uses the existing logical-line cursor context. At an empty
position its previous policy skips one trailing newline: one preceding logical
line preserves that behavior without copying the module prefix. Another newline
still ends the lookback. Local and module symbol selection is unchanged.

The work-count regression checks fewer than 100 cached-token reads near the end
of a 1,200-procedure module. It covers member-position rejection and identifier
lookup, plus continued/numbered/colon-led Call statements and the blank-line As,
operator and Call behavior. Existing casing and identifier suites also pass.

## Private component measurements

Controlled comparison bundles differ only in identifier cursor lookup. On a
969,412-character fixture derived from the private class, 60 warm checks at ten
varying offsets measured median/p95/max 2.14/4.30/9.88 ms before and
0.21/0.41/0.45 ms after. Thirty fresh-source checks, changing only a fixed-width
synthetic comment, measured 8.38/12.10/14.42 ms before and 6.56/9.44/10.75 ms
after. Private source and comparison bundles remain in ignored .vscode-test.
These measurements do not establish popup painting latency or absence of stalls.

## Opt-in renderer profiling

XLIDE_PERF_CPU_PROFILE=1 now profiles the renderer stress test in both the
extension host and renderer. Profiles and wall-clock request/response markers
are saved in its disposable workspace. The host records heartbeat delays over
20 ms; renderer samples include cycle start times. Both profiles are retained
on a failed stress run. Profile activation itself can add significant startup
overhead, so profiled timings are diagnostic and require an unprofiled repeat.


A repeated profiled run passed 256 fresh-source cycles and 16 hovers. Menu
median/p95/max changed from 111/151/184 ms to 111/133/140 ms, while maximum
recorded host heartbeat delay changed from 85.32 to 63.07 ms. Inclusive
identifier work changed from 5.38 to 3.72 seconds and casing from 10.19 to
8.10 seconds. Visible Backspace maxima were 50/77 ms and typing 51/48 ms;
these measurements do not claim a native editing speedup.

An intervening profiled run failed to observe Backspace at cycle 113. It lacked
caret/focus failure capture, so it is not classified as a product stall or
excluded as a test error. The probe now captures focus, synthetic-line suffix
booleans and actual document/caret metadata on failure; it prints no private
source or completion labels. Further sustained unprofiled coverage is required.


The full local suite passed 14,601 tests across 743 files with four workers.
The default-worker attempt hit two existing five-second storage-test timeouts;
no limits or tests were changed. An unprofiled repeat passed all 12 native/
workbook cases, 1,000 fresh-source cycles and 62 mouse hovers at 9fcfce85.
Visible Backspace median/p95/max was 16/37/61 ms; typing 17/46/61 ms; menus
120/156/651 ms; miss invalidation 1/18/77 ms. Mouse hover was 380/406/420 ms,
including Code's hover delay. The longest menu was cycle 573 after 30 ms idle.
Command-driven typing median/max was 2.14/14.98 ms and Backspace 2.02/85.21 ms;
the recorded host-heartbeat maximum was 48.72 ms. Declaration edit-to-hover was
15.29 ms and edit-to-code-actions 3.78 ms. Native deletion still worked during
an occupied host, including a forced stale cleanup context.

The cycle-113 failure did not repeat in this run, but remains unclassified.
The 651 ms popup outlier keeps the overall stalling/latency goal open. The
component optimization is validated; the remaining outlier needs a longer
correlated CPU capture rather than attribution from these unprofiled timings.
