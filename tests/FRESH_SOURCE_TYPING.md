# Fresh-source typing and hover validation

The actual-class renderer probe previously alternated two source texts,
`.cez` and `.ce`. That checks recovery but can reuse source-keyed caches after
the first cycle. Set `XLIDE_PERF_FRESH_SOURCES=1` to replace a fixed-width comment
nonce on the preceding synthetic statement between cycles. It changes only
the disposable probe procedure, so every subsequent member recovery starts
from different source text while the member expression stays identical.

The renderer verifies every nonce insertion and the extension document verifies
the final nonce. Escape precedes navigation so another provider's matching list
cannot intercept the movement keys. Mouse targets must be inside the editor's
visible scroll area before the hover is measured. The test writes a count-only
progress marker every 100 cycles and removes its markers on completion.

Compile, then run `tests/nativeBackspaceIntegration.config.mjs` with the private
workbook in `XLIDE_PERF_WORKBOOK` and `XLIDE_PERF_RENDERER_CYCLES=1000`. The
original workbook is copied before editing. The pinned Code version is 1.139.1;
`XLIDE_TEST_CODE_EXECUTABLE` can select the existing pinned installation.

## Results

On the automatic-recovery branch at 3e17f165, all 1,000 fresh-source cycles and
62 mouse hovers passed over seven minutes, with no listener warnings. Visible
Backspace median/p95/max: 16/34/62 ms. Typing: 17/46/62 ms. Menu recovery:
120/166/293 ms. Miss invalidation: 1/17/69 ms. Mouse hover median/max: 377/438 ms,
including the normal editor hover delay. These are DOM observations with
input/IPC/polling overhead, not GPU presentation times.

The largest menu outlier occurred at cycle 601 after 30 ms idle. Menu maxima
were 265/293/173/174 ms for 0/30/250/350 ms idle cadences. The built-in popup
reveal timer accounts for 100 ms of a new show, but the remaining cold/slow
work still needs profiling. These results establish sustained native editing,
recovery and resolved hovers; they do not establish nearly instantaneous menus
or absence of future stalls.

Earlier growing-comment variants were excluded: horizontal scrolling put a
hover target left of the editor viewport at cycle 64. A fixed-width nonce and
viewport hit-test validation avoid treating that test error as an editor stall.
Private source and profiles stay in ignored .vscode-test files.


After integrating main at 1b445e84 and simplifying viewport validation, all 12
workbook/native cases passed with 64 fresh-source cycles and four mouse hovers.
Visible Backspace and typing maxima were 64/68 ms; popup maximum was 316 ms.
Command-driven Backspace median/max were 5.69/132.40 ms, with a 103.17 ms host
heartbeat maximum. Fresh declaration hover returned updated content in 32.95 ms.
Those host-loop and popup outliers remain material and keep the latency goal
open for profiling; a green correctness run alone does not close it.
