# Smart Enter caret ownership

## Confirmed input race

Block insertion, comment continuation and With-member continuation previously
placed the caret immediately after `editor.edit` resolved, then retried later.
The retry checked ownership, but the initial placement did not. Native navigation
while an edit acknowledgement was pending could therefore paint correctly and
then be undone by the extension. With cases also opened suggestions on the
body line to which the caret had been returned.

The renderer baseline used a disposable VS Code 1.139.1 instance and an
independent debugger process. After the automatic edit painted, a test listener
held the extension host for 1.2 seconds while the acknowledgement was pending.
Native Up moved to the preceding line during that interval. After the host
resumed, all three paths moved back to the body. All three normal-placement
controls passed; all three navigation-ownership cases failed. This is a
correctness stress condition, not a claim about ordinary edit latency.

## Resulting behavior

A scoped selection observer now records explicit keyboard, mouse and command
navigation from before the edit request through deferred initial placement.
Text-edit adjustments have undefined selection kinds in the retained native
observations; native Up is a keyboard event and the previous rollback was a
command event. The deferred pass lets queued notifications arrive before making
its ownership decision.

Placement requires the original active editor/document, the single expected
edit revision, and the original or expected edit-adjusted caret. Navigation or
later typing relinquishes placement, including movement back to the original
position. The observer is disposed after placement or cancellation, and after
rejected or thrown edits. Dot suggestions are scheduled only after owned
placement and still verify the caret when their timer runs.

## Validation

The final native harness passed all nine cases: normal placement, native Up,
and Up followed immediately by typing for each of the three paths. Both
movement and typing painted during the deliberately busy host interval and
remained on the intended row afterward. Normal With cases retained their dot
suggestions; moving away left suggestions closed. Each case uses its own file
URI and waits for the intended visible caret before sending Enter.

Compilation and 181 focused tests passed, with two opt-in benchmarks skipped.
The focused suite includes 22 ownership/lifecycle regressions covering pending
navigation, keyboard/mouse/command notifications, later typing, rejected and
thrown edits, closing and switching editors, plus existing typing and editor
command regressions. Its document fixtures now model real edit revisions.

Run the native cases with `npx vscode-test --config
tests/nativeEnterIntegration.config.mjs`. The configuration pins VS Code and
allocates a localhost debugger port for the disposable instance. A supplied
`XLIDE_TEST_CODE_EXECUTABLE` can select an existing pinned installation. Normal
integration runs without this port skip the native suite.

This fix establishes ownership for the reproduced Enter race. It does not
attribute earlier synthetic wrong-row failures or every remaining large-class
latency sample to that race.
