# Shared Introduce Parameter binding index

The recursive binding path and external-caller filter previously created independent ProjectIndex instances for the same refactor. With recursion and an external owning-module call, the selected module was built twice. Qualified recursion also built other project modules twice. This duplicate construction is removed by supplying the existing lazy query-owned project to procedureCallBinding.

The selected module is still built only when binding is needed. Other modules are added once, only when a qualified recursive receiver or external caller requires them. The helper's standalone default keeps its existing standard-module construction behavior; the supplied project retains the refactor's actual module kinds. No cache survives a refactor query or mutates a source AST.

Deterministic guards exercise bare and module-qualified recursion plus 1, 100 and 1,000 owning-module external calls, alongside an unrelated same-name procedure. They require at most one build per needed module and independently verify complete primary and external outputs. All six added cases fail on parent 7eb1b086 because the primary module is built twice; the previous nine frozen-AST work controls pass. After the fix all 15 pass. Focused validation passes 99 tests across binding/work/receiver files; type checking passes.

Differential validation against 7eb1b086 covers 8,256 corpus sources and 19,814 complete call-site arrays/rendered sources, all identical, plus 216 generated full refactors, all identical and matching independently expected text. Lexer tokens/trivia and parser ASTs are frozen. Edit bounds and overlaps are checked explicitly; there are no uncaught exceptions. The generated ordinary cases do not by themselves establish all recursive receiver behavior; dedicated output/work tests supply that evidence.

## Measurement

Run node scripts/benchmark-call-site-arguments.mjs --recursive-refactor --project-callers --rounds=9, with --baseline=7eb1b086 for the parent. Fixtures contain 1,000 private module fields, 1/100/1,000 recursive calls and an equally sized module-qualified external call list, plus a foreign same-name call which must stay unchanged. Both versions produce correct complete edits in all 18 configurations.

Node v24.18.0, AMD Ryzen 7 9800X3D 8-Core Processor           ; three warmups and nine rounds, separate sequential processes in baseline/candidate/candidate/baseline order. Fresh source varies a leading comment, invalidating source-keyed caches; cached source is unchanged. Timing includes the pure refactor, with full edited-source verification outside the timed interval. Ranges below span the two process medians. Timings are mixed and do not establish an overall latency improvement; this change removes verified duplicate construction, not a measured universal speedup. No Office/editor latency is measured.

| Calls / recursive receiver / source | Before median range (ms) | After median range (ms) |
|---|---:|---:|
| 1/bare/cached | 1.359–1.413 | 0.955–1.239 |
| 1/bare/fresh | 2.226–2.252 | 2.040–2.741 |
| 1/Me/cached | 1.743–1.828 | 1.540–1.593 |
| 1/Me/fresh | 1.838–1.841 | 1.697–2.254 |
| 1/WithMe/cached | 1.313–1.377 | 1.147–1.157 |
| 1/WithMe/fresh | 1.874–1.907 | 1.727–1.781 |
| 100/bare/cached | 1.632–1.808 | 1.548–1.628 |
| 100/bare/fresh | 3.289–3.296 | 2.392–3.954 |
| 100/Me/cached | 2.013–2.034 | 2.515–3.890 |
| 100/Me/fresh | 2.824–2.975 | 2.735–4.326 |
| 100/WithMe/cached | 2.463–2.528 | 2.706–4.030 |
| 100/WithMe/fresh | 3.415–4.286 | 3.187–4.743 |
| 1000/bare/cached | 5.216–6.149 | 6.362–7.760 |
| 1000/bare/fresh | 6.968–7.482 | 7.236–8.286 |
| 1000/Me/cached | 6.144–6.358 | 5.986–6.459 |
| 1000/Me/fresh | 8.661–8.832 | 8.080–10.378 |
| 1000/WithMe/cached | 8.616–9.262 | 8.787–10.235 |
| 1000/WithMe/fresh | 10.859–12.790 | 10.817–16.367 |


Full repository validation is running before this change is marked ready.
