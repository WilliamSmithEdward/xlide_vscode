# Recovery after delayed native caret updates

Native Backspace can update the document before the extension host receives
its new caret. Recovery previously stopped waiting after one second. If host
work occupied the event loop longer, the initial probe still saw the old caret,
the expiry disposed the listener, and the later valid selection could not
reopen suggestions for `.ce` after deleting `z` from `.cez`.

A standalone public fixture, `ThisWorkbook.Sheets(1).cez`, reproduces this by
pausing the host for 1.5 seconds in a document-change listener after recovery
has installed its selection listener. Unchanged completion code fails on the
first native deletion with correct text, caret and renderer focus. A local
trace records the old caret, the host pause, the failed initial probe and
expiry before the queued selection can be observed.

Recovery now retains at most one pending wait per provider. A valid caret
completes it; a later text edit, navigation, editor switch, document close or
provider disposal cancels it. Empty dirty-state notifications preserve it.
Disposal also prevents asynchronous context loading from triggering recovery
later. There is no wall-clock expiry to race a delayed valid selection.

The native regression asserts that all six imposed pauses precede their caret
updates. All six recoveries pass. With foreground preconditions, deletion
painted in 4–30 ms and suggestions returned in 1,587–1,637 ms, including the
imposed pause. These measurements verify recovery under load, not a general
speed improvement. Its renderer artifacts use separate names from the
ordinary large-class capture.

## Validation

Compilation and type checking pass. The recovery suite adds seven unit cases
covering delayed selection, empty changes, superseding edits, close, switch,
bounded retained state and disposal during asynchronous loading. The former
expiry test now requires recovery from a late valid caret. All 15,534 unit
cases pass across 792 files, with 33 existing skips across seven files. All
39 editor integration cases pass.

All thirteen native integration cases pass both the 64-cycle ordinary run
and the 1,000-cycle profiled run. Each ordinary cycle uses real renderer keys
to delete `.cez` to `.ce`, requires a visible `Cells` suggestion, types `z`
and requires the missed-prefix menu to clear. A synthetic preceding statement
changes between cycles to prevent identical source reuse. Real mouse hovers
must resolve the synthetic variable's `Long` type every sixteen cycles.
The original large workbook is accessed through a disposable copy.

| Operation | 64 cycles: median / p95 / max, ms | 1,000 cycles: median / p95 / max, ms |
| --- | --- | --- |
| Backspace paint | 17 / 37 / 47 | 16 / 34 / 50 |
| Suggestion menu after deletion paint | 117 / 125 / 138 | 119 / 289 / 1,916 |
| Typing paint | 16 / 46 / 57 | 16 / 50 / 67 |
| Missed-prefix dismissal | 2 / 11 / 12 | 1 / 51 / 1,503 |
| Resolved mouse hover | 386 / 406 / 406 (4 hovers) | 379 / 413 / 419 (62 hovers) |

The long run uses pinned VS Code 1.139.1, host and renderer CPU profiling,
fresh-source edits and navigation observations, with word suggestions disabled.
The harness brings its owned page forward and verifies one focused editor
before each timed input sequence, hover and fresh-source sequence. It does
not restore focus inside timed polling. Focus setup is outside action timing
but contributes to the overall profiled workload. These runs are validation
observations, not a controlled before/after speed comparison.

Five profile/timing/observation files and the exact extension bundle and map
were retained locally. Their freshness, unchanged capture-time bundle hashes,
1,000 samples, 62 hovers and all seven artifact hashes were verified. Private
source, profiles and traces remain in ignored local artifacts.

## Limits and retained failures

An earlier ordinary candidate run had a focused menu failure at cycle 722,
after 722 successful recoveries and 45 resolved hovers. Native deletion still
painted while host frames included pauses over two seconds. That capture had
no recovery-state trace, so it cannot establish that this timer caused that
particular failure. The standalone reproduction confirms the timer defect
independently. Successful diagnostic runs with state logging do not erase
that failure.

Separate runs lost native window focus during hover after 32 cycles and during
fresh-source selection after 73 cycles. Both failed captures are retained;
neither establishes an extension stall. Foreground preconditions were added
to the owned test window before the completed long run.

Slow frames remain: the completed run's suggestion delay reached 1.9 seconds.
The profile also records host heartbeat delays and repeated canonical casing,
symbol and project-context work. Inclusive profile categories overlap; their
presence during a slow frame does not prove its cause. The broader latency
investigation remains open. This change preserves completion recovery when
host work delays the caret update.
