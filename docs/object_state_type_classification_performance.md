# Object-state type classification performance

Baseline: `60d5faa97cd35cffec740e68f97b5fb4b696b75a` (PR #1187). The baseline comparison restores objectState.ts; all surrounding files are identical.

The public object-state rule repeatedly resolves object types for function results, untouched module variables, local declarations and object arrays. Reuse one lazy resolver per public invocation across those paths and their procedure walks. The existing module-facts calculation also shares its resolver with returning-Nothing function classification. Direct Let-state queries that build a walk without the public rule create a resolver only after the existing walk-cache check. Resolver lifetimes remain local; source/module/symbol/activity/member-context cache ownership and freshness checks are unchanged. Private helpers receive the resolver rather than repeating unindexed classification. No dependency is added.

## Validation

Types and 110 focused tests pass. Fifty new tests include 21 public-rule work bounds over four families, repeated/distinct types and 10/100/1,000 procedures; 24 class/document/userform/UDT/enum/module eligibility tests assert complete exact findings; three New/Static/scalar exclusions; refreshed-context classification on a retained AST; and a direct Let-state query before the public rule fills the walk cache. The baseline fails 19 work bounds, passes two small work bounds and all 29 semantic/cache controls.

Full suite: 15,577 passed, 33 skipped across 793 passing files and seven skipped files, in 80.11 seconds. Complete outputs match 20,000 generated public-rule results, 2,000 generated full-module results and 24,768 corpus diagnostic/internal-error results over 8,256 sources with LF/CRLF/CR; zero differences. Generated cases cover parameter first-member reads, local/module objects, literal-index arrays and object-returning functions, metadata kinds/completeness/duplicates/implemented names, scalar/generic/project/host/unknown types, guards and Set/New/Nothing/error-handling/read variants. No diagnostic changes are intended.

## Work

Each fixture has N procedures and N+1 class surfaces. Functions return repeated or distinct project types; locals and arrays declare one corresponding type per procedure; module-variable fixtures declare one shared Class0 variable. The public rule independently produces no findings. Full-module expectations assert exact counts/messages/spans for missing-return warnings or unused declarations, and empty internal errors. Getter counters are untimed.

| N = 1,000 family | Types | Scope | Before class-name reads | After class-name reads |
| --- | --- | --- | ---: | ---: |
| functions | repeated | public-rule | 3,006,000 | 2,004 |
| functions | repeated | complete-module-diagnostics | 3,010,004 | 6,008 |
| functions | distinct | public-rule | 3,006,000 | 4,002 |
| functions | distinct | complete-module-diagnostics | 3,010,004 | 8,006 |
| module-variable | repeated | public-rule | 2,004,000 | 1,002 |
| module-variable | repeated | complete-module-diagnostics | 2,008,004 | 5,006 |
| locals | repeated | public-rule | 1,002,000 | 1,002 |
| locals | repeated | complete-module-diagnostics | 1,006,004 | 5,006 |
| locals | distinct | public-rule | 1,002,000 | 2,001 |
| locals | distinct | complete-module-diagnostics | 1,006,004 | 6,005 |
| arrays | repeated | public-rule | 2,004,000 | 1,002 |
| arrays | repeated | complete-module-diagnostics | 2,008,004 | 5,006 |
| arrays | distinct | public-rule | 2,004,000 | 2,001 |
| arrays | distinct | complete-module-diagnostics | 2,008,004 | 6,005 |

These fixtures bound classification work in the public object-state rule. Statement-level default-verdict queries and other analyzer consumers are outside this classification repair; no universal full-analyzer linear-work claim is made.

## Timings

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor. Sequential baseline/current/current/baseline trials, three warmups and nine measured rounds. Ranges are the two trial medians in milliseconds. Public-rule timing excludes parsing/symbol construction; complete-module timing includes analyzer setup and internal-error capture. Getter counters are replaced with plain properties before timing. Independent expectations and equality checks run outside the clock. No editor, cold-start or heap claim.

| Family | Types | Scope | N | Before median ms | After median ms |
| --- | --- | --- | ---: | ---: | ---: |
| functions | repeated | public-rule | 1 | 0.0241–0.0263 | 0.0218–0.0229 |
| functions | repeated | complete-module-diagnostics | 1 | 0.6000–0.7048 | 0.6068–0.6307 |
| functions | repeated | public-rule | 100 | 0.9287–1.3784 | 0.2777–0.3109 |
| functions | repeated | complete-module-diagnostics | 100 | 5.2591–6.3059 | 4.3983–4.5093 |
| functions | repeated | public-rule | 1000 | 52.3375–58.3420 | 2.5446–2.6345 |
| functions | repeated | complete-module-diagnostics | 1000 | 153.3960–166.0067 | 102.9048–106.3538 |
| functions | distinct | public-rule | 1 | 0.0025–0.0027 | 0.0029–0.0030 |
| functions | distinct | complete-module-diagnostics | 1 | 0.1750–0.2280 | 0.1751–0.1784 |
| functions | distinct | public-rule | 100 | 0.8570–0.9019 | 0.2585–0.2627 |
| functions | distinct | complete-module-diagnostics | 100 | 3.3487–3.5056 | 2.7463–2.7591 |
| functions | distinct | public-rule | 1000 | 55.7625–65.7978 | 3.4146–3.4202 |
| functions | distinct | complete-module-diagnostics | 1000 | 156.3917–161.7067 | 95.0173–99.8414 |
| module-variable | repeated | public-rule | 1 | 0.0036–0.0055 | 0.0034–0.0034 |
| module-variable | repeated | complete-module-diagnostics | 1 | 0.4049–0.6499 | 0.3795–0.3903 |
| module-variable | repeated | public-rule | 100 | 0.5772–0.5925 | 0.1904–0.1946 |
| module-variable | repeated | complete-module-diagnostics | 100 | 3.3254–3.5319 | 2.8125–3.0292 |
| module-variable | repeated | public-rule | 1000 | 36.5155–39.5062 | 2.3743–2.4100 |
| module-variable | repeated | complete-module-diagnostics | 1000 | 75.5527–80.6383 | 42.1498–43.8306 |
| locals | repeated | public-rule | 1 | 0.0030–0.0030 | 0.0031–0.0032 |
| locals | repeated | complete-module-diagnostics | 1 | 0.3564–0.3654 | 0.3571–0.4045 |
| locals | repeated | public-rule | 100 | 0.3726–0.3796 | 0.1858–0.1960 |
| locals | repeated | complete-module-diagnostics | 100 | 3.6105–3.7870 | 3.1892–3.8173 |
| locals | repeated | public-rule | 1000 | 20.0862–21.6672 | 3.6472–3.9538 |
| locals | repeated | complete-module-diagnostics | 1000 | 67.5404–71.5439 | 51.7516–52.8473 |
| locals | distinct | public-rule | 1 | 0.0015–0.0016 | 0.0017–0.0017 |
| locals | distinct | complete-module-diagnostics | 1 | 0.1821–0.2128 | 0.1860–0.1946 |
| locals | distinct | public-rule | 100 | 0.3273–0.3280 | 0.1658–0.1870 |
| locals | distinct | complete-module-diagnostics | 100 | 2.5972–2.7792 | 2.5444–2.5797 |
| locals | distinct | public-rule | 1000 | 21.8310–23.7642 | 4.4151–4.4600 |
| locals | distinct | complete-module-diagnostics | 1000 | 64.7679–71.4034 | 46.8075–48.5316 |
| arrays | repeated | public-rule | 1 | 0.0033–0.0034 | 0.0032–0.0032 |
| arrays | repeated | complete-module-diagnostics | 1 | 0.2930–0.2930 | 0.2775–0.2944 |
| arrays | repeated | public-rule | 100 | 0.5151–0.5263 | 0.1333–0.1362 |
| arrays | repeated | complete-module-diagnostics | 100 | 3.1942–3.3797 | 2.7557–2.8346 |
| arrays | repeated | public-rule | 1000 | 37.0804–39.6157 | 3.9251–4.0535 |
| arrays | repeated | complete-module-diagnostics | 1000 | 95.7685–96.1874 | 56.0881–58.4828 |
| arrays | distinct | public-rule | 1 | 0.0018–0.0019 | 0.0017–0.0017 |
| arrays | distinct | complete-module-diagnostics | 1 | 0.1753–0.1760 | 0.1790–0.2135 |
| arrays | distinct | public-rule | 100 | 0.5156–0.5365 | 0.1553–0.1566 |
| arrays | distinct | complete-module-diagnostics | 100 | 2.7614–2.9188 | 2.5473–2.5958 |
| arrays | distinct | public-rule | 1000 | 41.1550–44.0814 | 4.1966–4.3477 |
| arrays | distinct | complete-module-diagnostics | 1000 | 88.2819–90.7116 | 55.3665–55.7418 |

The one-procedure full-module ranges overlap for functions, locals and repeated arrays. Distinct one-procedure arrays are slower in both full-module trials, and several one-procedure public-rule fixtures are also slower. Repeated 100-procedure local full-module results overlap and include a slower after trial. All 1,000-procedure public-rule and full-module fixtures improve in both trials; do not infer a universal small-case gain.

Reproduce from the repository root, sequentially:

```powershell
node scripts/benchmark-object-state-type-classification.mjs --baseline=60d5faa97cd35cffec740e68f97b5fb4b696b75a
node scripts/benchmark-object-state-type-classification.mjs
node scripts/benchmark-object-state-type-classification.mjs
node scripts/benchmark-object-state-type-classification.mjs --baseline=60d5faa97cd35cffec740e68f97b5fb4b696b75a
```
