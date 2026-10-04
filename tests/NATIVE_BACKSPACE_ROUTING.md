# Native Backspace routing

The VBA Backspace keybinding previously routed every ordinary deletion through
`xlide.vba.smartBackspace` in the extension host. A busy or stuck host therefore
prevented Backspace while native typing, Enter and arrows could still work.

The keybinding now requires a cleanup context: one caret on a whitespace-only
indented line, or an empty continued-comment marker. Ordinary code, selection
and multiple-caret deletion use VS Code's native handler. The context reads only
the caret line and, for a comment marker, its preceding line, and sends
`setContext` only when its boolean changes.

A true cleanup flag can become stale while the host is busy: the user can type
ordinary code into the formerly blank line before the host updates that flag.
The cleanup binding therefore runs native `deleteLeft` first through VS Code's
`runCommands`, then invokes `xlide.vba.finishBackspaceCleanup`. A busy host can
delay optional cleanup but cannot hold up the initial deletion, even with a
stale true flag. The follow-up validates the original line/caret, exact native
deletion, document version and current editor/caret. It discards cleanup after
another edit, navigation, close or a rejected edit, and never retries deletion.
Native deletion and any remaining indentation/comment cleanup share one Undo.

The event tracker retains two line/caret snapshots because native selection
and text notifications can arrive in either order. Dirty-state notifications
with no text changes preserve pending cleanup until the native caret arrives.
It introduces no full-module reads, timers or response-path sleeps.

## Verification

Run `npm run compile`, then
`npm exec -- vscode-test --config tests/nativeBackspaceIntegration.config.mjs`.
The configuration pins VS Code 1.139.1, chooses an unused loopback debugger port,
and opens only a disposable workspace/profile. Set `XLIDE_TEST_CODE_EXECUTABLE`
to an existing installation's executable to avoid downloading that pinned build.
Set `XLIDE_PERF_WORKBOOK` to the private workbook to include its optional probes.
The original workbook is copied before editing; no source or CPU profiles are
committed.

A separate Node process sends Backspace through Chromium's renderer input API
while the extension host is deliberately blocked for 1.2 seconds. The native
route must change the visible editor line during that interval. Both ordinary
false-context and deliberately stale true-context routes must pass. A third
probe starts on an indented blank line, types code while the host is blocked,
then presses Backspace while the true flag remains stale. Physical-key cleanup
probes check indentation and continued comments, with one Undo restoring the
entire action. This distinguishes native routing from invoking `deleteLeft`
through the extension host.

The actual-class renderer probe defaults to 24 `.cez` -> `.ce` -> `.cez` cycles at
0/30/250/350 ms idle cadences. Every Backspace must recover a visible Cells row;
the subsequent miss must hide it before another recovery can count. It records
visible DOM updates, suggestion visibility and resolved mouse hovers. Set
`XLIDE_PERF_RENDERER_CYCLES=200` for a longer session; a mouse hover is checked
every 16 cycles. Measurements include debugger roundtrip
and polling overhead, rather than GPU presentation timing. Provider/command
measurements are reported separately and may still wait for the host.

Unit regressions cover the manifest gate, context event transitions, close,
selection/multiple-caret routing, and zero whole-module reads/context traffic
for ordinary typing. Existing continued-comment/whole-indent integration
checks preserve those behaviors.

## Earlier results before native-first cleanup

After integrating main at 49ed5c90, three VS Code 1.139.1 runs passed all eight
cases. In 72 actual-class renderer cycles, visible Backspace medians were
14/16/9 ms and maxima were 32/33/30 ms. Typing medians were 30/29/30 ms and maxima
46/32/45 ms. Every recovery displayed Cells and every `.cez` miss cleared it.
The first visible recovery in each run took an additional 136/124/133 ms after
deletion; warm menu-update medians were 6/7/9 ms. These first-widget samples are
reported separately and remain worth investigating.

The deliberately blocked-host native probes updated the visible line in
31/32/31 ms during the 1.2-second stall. Each extension-bound control remained
unchanged throughout its 606-607 ms observation window, then deleted after the
host resumed. The earlier isolated probe measured 13 ms on the native route.

All 144 command-driven typing/Backspace pairs restored the actual class, and
24 command-driven member recovery/hover cycles returned Cells and a hover.
Typical command Backspace was 1.81/1.74/2.06 ms, with maxima 107.49/85.12/109.09 ms;
maximum host heartbeat delays were 61.60/53.10/65.21 ms. Fresh declaration hover
returned updated content in 16.57/31.96/31.37 ms. Command outliers and the first
menu show remain distinct from successful native editing while the host stalls.
The intermittent latency goal and issues #964/#985 remain open.

The merged-base full unit suite passed 14,474 tests across 729 files, with
32 tests and seven files intentionally skipped. Compilation and the 34 focused
routing/lifecycle/cleanup tests passed.


## Stale-context reproduction and correction

On the previous narrow-context binding, typing out of an indented line while
blocking the host left Backspace unchanged for its entire 604 ms observation
window. With native-first cleanup, the same transition deleted in 7 ms while
the 1.2-second host block was still in progress. An initial synthetic run also
passed ordinary and forced-stale routing in 32 ms each. Physical indent and
continued-comment cleanup took 50–66 ms and each restored fully with one Undo.
These values include renderer polling/roundtrips; they are not GPU timing.

Focused regressions additionally cover both notification orders, interleaved
no-change notifications, intervening typing, navigation, focus/version changes,
close, multiple carets, selections, Undo/Redo and rejected cleanup.


The 200-cycle private-class run passed all 12 cases. Visible Backspace had a
16 ms median, 33 ms p95 and 47 ms maximum; typing had a 30 ms median, 39 ms p95
and 47 ms maximum. Warm menu recovery had a 5 ms median (first show 118 ms).
All 12 mouse hovers resolved in 371–412 ms with the editor's normal hover delay;
none stayed at Loading. Fresh declaration hover returned updated content in
34.75 ms. Forced-stale and transition Backspace still deleted during the busy
interval in 19/31 ms. The long run also emitted renderer listener-leak warnings
from VS Code's suggestion-menu status code. That warning is recorded for
further investigation; these measurements do not establish its cause or an
absence of future long-session stalls.
