# Tree tab events with no closures

ExplorerFollow handled every TabChangeEvent by asking which modules lost their last tab. Its production callback flattens all tab groups, then modulesWithNoTabLeft enumerates each tab URI into a Set. Active-state changes and new-tab events have no closed tabs, so this entire traversal cannot produce a closed module. VS Code's installed event contract explicitly includes active-state changes in changed tabs.

Return before calling the closure resolver when event.closed is empty. The active-editor/caret events still schedule tree following, and real closure events keep their existing module, custom editor and diff semantics. No new cache or helper is added.

## Reproduction and boundaries

```powershell
node scripts/benchmark-tree-tab-events.mjs --baseline=88ca7c34
node scripts/benchmark-tree-tab-events.mjs
```

The benchmark bundles the actual controller, extracts its production modulesClosedBy callback from extension.ts and the actual tabUris/modulesWithNoTabLeft function declarations from vbaDocumentLocation.ts. A small event host supplies real-shaped Tab objects whose input is TabInputText. Document ownership decoding is stubbed to return undefined, so the actual-close timing control closes a non-module text tab. Real module/custom/diff closure semantics are covered by the focused tests, not this stub. The harness asserts that every enumerated text tab contributes one URI conversion, and that nonempty closure events still invoke the resolver.

For each sample, 200 events are emitted over 1, 50 or 1,000 open text tabs across up to four groups. Changed and opened events carry one tab and no closures; empty events have no entries; the nonmodule-close control carries one closed text tab. Preparation, result checks and disposal are outside timing. Callback/group/URI counters run in a separate untimed pass. Three warmups precede 15 measured samples; before and after run sequentially after the test suite, without other audit workloads. This is isolated event/resolver work, excluding real VS Code UI, editor activation, caret tracking, backend access and tree reveals.

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor; measured locally 2026-10-03. Times are milliseconds per 200-event batch. Values rounded to zero are shown as less than 0.001, not zero execution cost.

| Fixture / open tabs | URI conversions before | URI conversions after | Median before | Median after | p95 before | p95 after |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| changed-1 | 200 | 0 | 0.109 | 0.003 | 0.351 | 0.015 |
| changed-50 | 10,000 | 0 | 1.071 | 0.002 | 1.297 | 0.002 |
| changed-1000 | 200,000 | 0 | 19.603 | 0.002 | 21.285 | 0.016 |
| opened-1 | 200 | 0 | 0.080 | 0.002 | 0.138 | 0.002 |
| opened-50 | 10,000 | 0 | 0.981 | 0.002 | 1.153 | 0.013 |
| opened-1000 | 200,000 | 0 | 18.787 | <0.001 | 20.486 | 0.002 |
| empty-1 | 200 | 0 | 0.065 | <0.001 | 0.089 | 0.001 |
| empty-50 | 10,000 | 0 | 0.966 | <0.001 | 1.221 | 0.001 |
| empty-1000 | 200,000 | 0 | 18.896 | <0.001 | 21.901 | 0.001 |
| nonmodule-close-1 | 200 | 200 | 0.096 | 0.101 | 0.105 | 0.108 |
| nonmodule-close-50 | 10,000 | 10,000 | 1.031 | 1.060 | 1.220 | 1.382 |
| nonmodule-close-1000 | 200,000 | 200,000 | 18.915 | 18.965 | 19.547 | 19.329 |

Every no-close fixture removes 200 resolver calls and tab-group enumerations. With 50 open tabs it also removes 10,000 URI conversions; 1,000 tabs removes 200,000. Actual-close controls retain these costs and show small mixed timing changes. The result establishes avoided work on the no-close path, not a universal interaction latency improvement. Closing real tabs is a separate path still under audit.

## Validation

Three new regressions emit 200 changed/opened/empty events each and assert no resolver enumeration, module folding or extra reveal. All three fail before the guard. Existing simulated-close events now contain a closed Tab, matching the VS Code contract, and their late-reveal folding and following-the-remaining-editor assertions remain.

Type check passed. Focused caret/follow/tree/folder/closing-tab suites: 116 tests across five files. Full suite: 628 files, 12,872 passed, 13 skipped. Real VS Code integration was not run. No dependencies added. Broader expand/collapse, child-list redraw and repository audits remain incomplete.
