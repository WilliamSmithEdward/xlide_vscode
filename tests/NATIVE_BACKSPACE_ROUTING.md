# Native Backspace routing

The VBA Backspace keybinding previously routed every ordinary deletion through
`xlide.vba.smartBackspace` in the extension host. A busy or stuck host therefore
prevented Backspace while native typing, Enter and arrows could still work.

The keybinding now requires a cleanup context: one caret on a whitespace-only
indented line, or an empty continued-comment marker. Ordinary code, selection
and multiple-caret deletion use VS Code's native handler. The context reads only
the caret line and, for a comment marker, its preceding line, and sends
`setContext` only when its boolean changes. Cleanup still needs a responsive
extension host; the context is updated by editor events in that host.

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
route must change the visible editor line during that interval. A positive
control enables the extension binding and must stay unchanged during the same
interval, then delete once the host resumes. This directly distinguishes native
routing from merely invoking `deleteLeft` via the extension host.

The actual-class renderer probe performs 24 `.cez` -> `.ce` -> `.cez` cycles at
0/30/250/350 ms idle cadences. Every Backspace must recover a visible Cells row;
the subsequent miss must hide it before another recovery can count. It records
visible DOM updates and suggestion visibility, including debugger roundtrip
and polling overhead, rather than GPU presentation timing. Provider/command
measurements are reported separately and may still wait for the host.

Unit regressions cover the manifest gate, context event transitions, close,
selection/multiple-caret routing, and zero whole-module reads/context traffic
for ordinary typing. Existing continued-comment/whole-indent integration
checks preserve those behaviors.
