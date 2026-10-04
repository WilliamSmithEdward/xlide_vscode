# Module object parameter type performance

Baseline: `bb014ca7b4949456cbbb0a5a600a317b811195c4` (PR #1182). Baseline comparisons restore objectState.ts; the surrounding production sources stay identical between variants.

The module object-state facts pass classified each eligible procedure parameter by scanning the supplied project surfaces. That pass runs once per module/context, but N parameters over N project classes still caused quadratic name work. Build one lazy object-type resolver within the existing facts calculation and reuse it for parameter classification. The three-line runtime change preserves parameter-array/accessor/activity guards, known-object eligibility, ambiguity and first-member-read facts. The existing source/module/activity/context cache contract stays intact; the resolver exists only during a facts calculation. No dependency is added.

## Validation

Types and 55 focused tests pass. Twenty new tests cover six full-module work bounds at 10/100/1,000 procedures with repeated/distinct types, six project-surface kinds, seven generic/host/project/unknown/scalar/case type controls, and new-context freshness under a retained module AST. Controls check exact diagnostic messages/spans for an unset local passed to a parameter whose first use reads a member. The baseline fails all six work bounds and passes all 14 semantic controls.

Full suite: 15,492 tests passed, 33 skipped across 790 passing files and seven skipped files, in 77.57 seconds. Complete results match 20,000 generated public object-state-rule outputs, 2,000 generated full-module outputs and 24,768 corpus diagnostic/error outputs across 8,256 sources with LF/CRLF/CR; zero differences. Generated cases vary class kind/completeness/duplicates, type aliases, guards, preceding statements, Set state and error handling. No diagnostic changes are intended.

## Work and timings

Each fixture has N procedures with one object parameter each and N+1 project classes. Types repeat Class0 or visit N distinct classes. Empty procedures independently produce no diagnostics/errors, isolating parameter-fact classification while measuring complete-module calls. Untimed name getters count the whole analyzer. Getters are removed before all timing; metadata is plain objects and arrays.

| N = 1,000 fixture | Before class-name reads | After class-name reads |
| --- | ---: | ---: |
| Repeated type | 1,006,004 | 5,006 |
| Distinct types | 1,006,004 | 6,005 |

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor           . Sequential baseline/current/current/baseline trials, three warmups and nine measured rounds. Complete-module setup and internal-error capture are inside the clock; independent expected-empty diagnostics/errors and equality assertions are outside it. Ranges give two trial medians in milliseconds. No editor, cold-start or heap claim.

| Types | Procedures | Before median ms | After median ms |
| --- | ---: | ---: | ---: |
| Repeated | 1 | 0.6205–0.7050 | 0.6110–0.6467 |
| Repeated | 100 | 4.4039–4.5491 | 4.1358–5.3065 |
| Repeated | 1000 | 66.5888–72.5246 | 47.3497–48.0431 |
| Distinct | 1 | 0.1838–0.1973 | 0.1941–0.1970 |
| Distinct | 100 | 2.4152–2.4154 | 2.2173–2.3203 |
| Distinct | 1000 | 64.4540–67.0115 | 43.9572–47.0587 |

Both one-procedure trial ranges overlap. Repeated-type 100-procedure ranges overlap and one after trial is slower. The first project lookup builds an eligible-surface index, so isolated small queries need not improve. Other many-procedure consumers still perform default-member/type queries; this fix addresses the attributed parameter-facts scan only. Those residual consumers remain in the audit.

Reproduce from the repository root:

```powershell
node scripts/benchmark-module-object-parameter-types.mjs --baseline=bb014ca7b4949456cbbb0a5a600a317b811195c4
node scripts/benchmark-module-object-parameter-types.mjs
```
