# Readonly helper input copies

Baseline: `b09c9682e3876b3ff6cb11536e6a7e1f006566ca` (published PR #1203 merge). Measurements used source head `ae121b1218af034d52cedf8e2ca42f725db10a07`; its runtime sources match the published merge. Only its report baseline pin changed after rebasing. Baseline builds restore all ten changed analyzer source files. Surrounding runtime files are identical.

Remove 31 redundant token-array spreads before matchParenFrom/splitTopLevelTokenGroups. Both helpers accept readonly arrays and only read them by index; the splitter creates independent output group arrays. Token identities, spans and call boundaries are unchanged. Keep the existing filtered-token semantics, including the separate comment-filter candidate. No new cache, ownership or lifetime is introduced.

Changed production sites:

- src/analyzer/diagnostics/typeInference.ts: 5
- src/analyzer/diagnostics/straightLineValues.ts: 1
- src/analyzer/diagnostics/rules/typeMembers.ts: 1
- src/analyzer/diagnostics/rules/overflow.ts: 2
- src/analyzer/diagnostics/rules/lateBoundMembers.ts: 8
- src/analyzer/diagnostics/rules/handlerFlow.ts: 4
- src/analyzer/diagnostics/rules/formContents.ts: 4
- src/analyzer/diagnostics/rules/errorValues.ts: 3
- src/analyzer/diagnostics/rules/collectionState.ts: 1
- src/analyzer/diagnostics/rules/accessData.ts: 2

## Validation

Types and 133 focused tests across 12 files pass. Five new cases cover three public allocation-work bounds at 10/100/1,000 extra call arguments and two frozen shared-input cases, through both public and complete diagnostics. Baseline fails all three work checks and passes both frozen-input controls. Public work checks independently assert the entire exact ProgID error code/text/span. Frozen cases preserve complete expected diagnostic output, including the existing unused-variable warning, and require no internal errors or input mutation.

Full suite: 15,807 passed, 33 skipped across 801 passing files and seven skipped files, in 96.06 seconds.

Complete outputs match 20,000 generated public rule results, 2,000 generated metadata modules, 10,000 generated call/grouping modules and 24,768 corpus runs over 8,256 sources with LF/CRLF/CR; zero differences. Calls cover ProgID/RegExp/Collection/control-add/grouping/error values, named and nested arguments, alongside host/form/conditional/With metadata cases. Full comparisons include every diagnostic and internal error. Focused existing suites exercise Access, collection state, error values, forms, handler flow, type members and straight-line contexts.

## Work and allocation scope

One wide CreateObject call has N extra arguments; the repeated-call fixture has N ordinary one-argument calls. Instrument each relevant cached statement array's iterator only outside timing, restore its original descriptor in finally, and assert exact independent public output. The unchanged namesIn walk still iterates the source array once. The two helper-input spreads previously iterated it twice more and allocated two temporary arrays. At 1,000 extra arguments this eliminates 4,014 copied token references. Counts measure array iteration/copies on this public path, not heap bytes or every removed call site.

| Fixture | N | Source token elements | Before iterators | After iterators | Before yielded elements | After yielded elements |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| wide-call | 1 | 9 | 3 | 1 | 27 | 9 |
| wide-call | 100 | 207 | 3 | 1 | 621 | 207 |
| wide-call | 1000 | 2,007 | 3 | 1 | 6,021 | 2,007 |
| many-calls | 1 | 7 | 3 | 1 | 21 | 7 |
| many-calls | 100 | 700 | 300 | 100 | 2,100 | 700 |
| many-calls | 1000 | 7,000 | 3000 | 1000 | 21,000 | 7,000 |

## Timings

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor. Sequential baseline/current/current/baseline trials, three warmups and nine measured rounds. Ranges are the two trial medians in milliseconds. Timed arrays use normal iteration; instrumentation is removed before timing. Public parsing/symbol setup is excluded; complete-module setup is included. Assertions run outside the clock. Exact public/primary full findings and complete full snapshots are checked; all four trials have identical full-output hashes. No editor/heap/cold-start or universal speedup claim.

| Fixture | Scope | N | Before median ms | After median ms |
| --- | --- | ---: | ---: | ---: |
| wide-call | public-rule | 1 | 0.0869–0.0886 | 0.0831–0.0865 |
| wide-call | complete-module-diagnostics | 1 | 1.2377–1.3574 | 1.3864–1.4204 |
| wide-call | public-rule | 100 | 0.1250–0.1259 | 0.1015–0.1268 |
| wide-call | complete-module-diagnostics | 100 | 2.0191–2.0808 | 1.8996–2.0117 |
| wide-call | public-rule | 1000 | 0.5242–0.5345 | 0.4656–0.4693 |
| wide-call | complete-module-diagnostics | 1000 | 9.3465–10.4182 | 9.2538–9.2768 |
| many-calls | public-rule | 1 | 0.0824–0.1079 | 0.0444–0.0445 |
| many-calls | complete-module-diagnostics | 1 | 0.7601–1.0184 | 0.7487–0.7833 |
| many-calls | public-rule | 100 | 0.3014–0.3040 | 0.2732–0.2777 |
| many-calls | complete-module-diagnostics | 100 | 7.7708–8.1284 | 7.4438–8.0101 |
| many-calls | public-rule | 1000 | 2.6071–2.7174 | 2.3571–2.4003 |
| many-calls | complete-module-diagnostics | 1000 | 70.9678–75.4656 | 72.4268–74.5836 |

Both 1,000-unit public-rule fixtures improve modestly in both trials. Complete diagnostics for 1,000 ordinary calls have overlapping/mixed medians; no elapsed-time improvement is established there. Single wide-call complete diagnostics are slower in both trials. Several other small/full ranges overlap. The confirmed benefit is fewer temporary arrays and copied token references; these data do not establish a universal or editor-latency gain.

Reproduce sequentially from the repository root:

```powershell
node scripts/benchmark-readonly-helper-copies.mjs --baseline=b09c9682e3876b3ff6cb11536e6a7e1f006566ca
node scripts/benchmark-readonly-helper-copies.mjs
node scripts/benchmark-readonly-helper-copies.mjs
node scripts/benchmark-readonly-helper-copies.mjs --baseline=b09c9682e3876b3ff6cb11536e6a7e1f006566ca
```
