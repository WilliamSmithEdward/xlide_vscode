# Automatic member completion recovery

Keyboard Backspace recovery previously invoked `editor.action.triggerSuggest`
without options, making it a manual request. With word suggestions disabled,
`.cez` after a successful recovery left a visible **No suggestions** widget
instead of dismissing the list. Repeated recovery cycles also emitted listener
warnings from the installed VS Code suggestion-status menu implementation.

Recovery now passes `{ auto: true }`. Explicit user-triggered suggestions keep
their normal behavior. The option is supported at the extension's minimum
VS Code version: see [TriggerSuggestAction in VS Code 1.95.0](https://github.com/microsoft/vscode/blob/1.95.0/src/vs/editor/contrib/suggest/browser/suggestController.ts).

The optional actual-class renderer harness checks visible Cells recovery and
clearing that row on each miss. `XLIDE_PERF_RENDERER_CYCLES=200` lengthens the
session and checks a resolved mouse hover every 16 cycles. To isolate truly
empty member results, set `XLIDE_PERF_WORD_SUGGESTIONS=0`: it temporarily disables
word suggestions in the disposable workspace's language settings and requires
the whole miss widget to dismiss. The previous override is restored afterward.
Normal word-suggestion mode can have other matching rows, so that mode requires
Cells invalidation and dismissal of empty/loading status messages, without
assuming the whole list is empty. Failure diagnostics
report widget classes, its standard status message, row count and icon kinds,
without printing completion labels or private source.

## Measurements

The prior command failed the isolated miss-dismissal check on cycle 0, showing
**No suggestions** throughout the four-second observation window. With automatic
recovery, all 200 cycles and 12 mouse hovers passed with no listener warnings.
Visible Backspace: median 17 ms, p95 33 ms, max 47 ms. Typing: median 16 ms,
p95 46 ms, max 62 ms. Popup recovery: median 118 ms, p95 136 ms, max 213 ms.
Hovers resolved in 371–419 ms with the normal editor hover delay.

These are DOM observations including input/IPC/polling overhead, not GPU timing.
Dismissing an empty list causes the next popup to incur VS Code's own 100 ms
reveal timer, verified in the pinned installation. The previous manual list could
remain open with a message and update its rows in about 5 ms, so these runs are
not an overall popup-latency speedup claim. They establish correct miss dismissal
and successful recovery; remaining cold/slow samples stay visible in the report.


The normal-profile 200-cycle run also passed, with 12 resolved mouse hovers and
no listener warnings. Its `wordBasedSuggestions` value was
`offWithInlineSuggestions`. Visible Backspace: median 17 ms, p95 32 ms, max 46 ms;
typing: median 19 ms, p95 46 ms, max 47 ms; popup recovery: median 118 ms,
p95 134 ms, max 156 ms. Hovers resolved in 352–414 ms. The earlier diagnostic
that required an entirely hidden widget was invalid when another provider had
matching rows; it is not counted as a failure of member recovery.
