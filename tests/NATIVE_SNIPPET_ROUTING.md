# Native snippet arrow routing

While a snippet was active, up/down/left/right invoked the extension-host
leaveSnippetAndCursorMove command. It awaited the workbench leaveSnippet command
before sending cursorMove back to the editor. Heavy analysis could therefore
stall navigation, and typing during that interval could edit the previous row.

The bindings now run both workbench commands in a renderer runCommands sequence,
with the same direction, line/character step and suggestion-widget guard. The
public legacy command remains available with its explicit-call lifecycle guards.

## Busy-host regression

The opt-in runner tests/nativeSnippetIntegration.config.mjs pins Code 1.139.1
and creates its own workspace, user data and localhost debugger port. A separate
process sends physical keys and observes visible caret geometry while the
extension host is deliberately occupied for 1.2 seconds. The fixture is a real
SnippetString placeholder with the caret inside it; the control is the same
editor without a snippet. No suggestion widget is visible.

The original bindings failed all four snippet directions: no movement was
observed during the 600 ms window (606/605/607/607 ms elapsed). Ordinary controls
painted movement in 32/20/32/7 ms. Final caret checks verified one movement after
the host resumed. The extended committed probe additionally failed arrow-then-
typing in snippet mode: text appeared on the previous row. Its ordinary control
passed. The complete baseline had five passing controls and five failures.

With native bindings, the first comparison passed all eight direction/control
cases; snippet up/down/left/right painted in 34/17/35/16 ms and exited the
snippet. The committed ten-case probe also passed: snippet directions painted
in 15/32/32/32 ms, and immediate arrow-then-typing painted the intended row in
17 ms, leaving the previous row unchanged. These include debugger round trips
and DOM polling, not isolated GPU paint time. They are busy-host routing
checks, not a claim that all large-class latency is resolved.

Manifest consistency checks recognize string and command-object entries in
runCommands, and validate their command IDs. Four routing regressions prevent
arrow bindings from depending on the extension-host command again. Existing
explicit-call lifecycle checks remain in place. Compilation and 22 targeted
manifest/lifecycle unit tests passed.

The earlier large-class synthetic edit failures did not establish snippet
mode at their failure points; this proven stall is not claimed as their cause.

After rebasing onto main including #1123, compilation, all ten native cases,
all 25 Completion editor surface cases and the 22 targeted unit tests passed.
