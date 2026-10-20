# Immediate completion response scheduling

Completion used a 150 ms project-context budget after its type/member/event
fast paths. Bare identifiers, keywords and labels waited on that build even
when current-module candidates were available. Each edit invalidates the
project editor cache, so the wait could recur while typing.

Completion now returns the current synchronous context without awaiting a
project build. Uncached lists are incomplete; the next request can incorporate
the warmed project context. Scheduled warming also serves loose modules,
starts on the next timer turn, is deduplicated by document/version, and is
discarded after edits, close or project invalidation. Project loading cannot
begin synchronously in front of this response.

Cross-module candidates may require a subsequent request while the cache warms.
No suggestion widget is forced open. Source lexing, local-context construction,
other extension-host work, VS Code transport and UI painting remain relevant.

## Validation and measurements

The fake-clock regression holds project completion for 150 ms. Local completion
must return at time zero, include the current local variable, be incomplete and
never invoke the budgeted project build. Another test verifies that the next
incomplete-list request includes warmed cross-module procedures and becomes
complete. Scheduling tests cover deferred startup, deduplication, edits, close
and invalidation.

The real VS Code 1.139.1 suite `Immediate completion surfaces` loads a
1,201-procedure module, types `instantvalue` character by character and requests
completion after every actual edit. It also types `name` after
`ThisWorkbook.Sheets(1).` and verifies later cross-module candidates.

An isolated run with the project loaded first measured edit-to-provider-command
result latency over 12 characters:

- Median: 13.2265 ms.
- Slowest sample: 41.3190 ms.

These include editor edits and command transport, but do not measure suggestion
widget painting. They are descriptive samples, not CI thresholds.

A fresh-start run still showed a 121.5397 ms slowest sample. A run concurrent
with the full unit suite showed a 298.6757 ms slowest sample. Removing the
budgeted wait does not establish nearly instantaneous latency for every cold
or busy-host request; those spikes remain an investigation target.

```powershell
npm run compile
npx vscode-test --code-version 1.139.1 --grep 'Immediate completion surfaces'
npx vitest run tests/completionResponseScheduling.test.ts
```
