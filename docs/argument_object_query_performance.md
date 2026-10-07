# Argument object query performance

Baseline: `ba3c4e5228c4c1a0b25c8b42dab8ba57b49ae37e` (PR #1180). Both argumentTypes.ts and typeInference.ts are restored for baseline comparisons.

The argument rule repeatedly resolved expected object types and checked implemented-name lists per argument. Create one lazy type, interface-sharing and membership query bundle per public argument-rule invocation and pass it through every statement, expression and member-call validator. Object-type gates and compatibility checks reuse that bundle across procedures. Held-object facts remain statement-specific. The optional helper parameters preserve the existing unindexed path when callers omit the bundle; each new public rule invocation observes fresh metadata. No dependency is added.

## Validation

Types and 128 focused tests pass. Twenty new tests cover nine full-module work bounds at 10/100/1,000 arguments, three isolated public-rule cross-procedure bounds, three incompatible-value/span/freshness controls, and five statement/Call/expression/member-expression/member-statement forms. The baseline fails 15 work tests and passes five: three semantic/freshness controls plus the two smallest declared/held work cases.

Full suite: 15,474 tests passed, 33 skipped across 789 passing files and seven skipped files, in 88.67 seconds. Complete output comparisons match 20,000 generated public argument-rule results, 2,000 generated full-module results and 24,768 corpus diagnostic/error results over 8,256 sources with LF/CRLF/CR; zero differences. Generated cases vary class kinds, duplicate eligibility, completeness, implemented-name aliases, host/project/generic/scalar/unknown types, ByRef/ByVal, New/Nothing/scalar/ActiveSheet/arithmetic/held arguments and all five call forms. No diagnostic changes are intended.

## Work

Each fixture has N calls and N extra classes. The Class2 implemented-name list has N entries with Class1 at its end. Declared, held-object and New arguments independently produce no diagnostics or internal errors. Counters include full-module work and run only outside timing.

| N = 1,000 path | Before class-name reads | After class-name reads | Before interface-string reads | After interface-string reads |
| --- | ---: | ---: | ---: | ---: |
| declared | 1,011,026 | 9,029 | 0 | 0 |
| held | 1,012,023 | 7,018 | 1,000 | 1,000 |
| new | 3,015,014 | 7,018 | 1,000,000 | 1,000 |

The declared argument's unproven runtime class stays quiet, preserving the prior behavior. Held-object list reads were already linear in this fixture; its repeated class-name scans are removed. Across many procedures, other full-analyzer consumers still scan project metadata: at 100 procedures, declared and New fixtures retain 31,916 and 10,915 full-module name reads. Therefore the cross-procedure bound specifically measures the public argument rule, not the entire analyzer. That residual is tracked separately.

## Timings

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor           . Sequential baseline/current/current/baseline runs, three warmups and nine measured rounds. These are warm complete-module diagnostic calls including analyzer setup and internal-error capture. All getters and index proxies are removed before timing; lists and metadata are plain arrays/objects. Independent expected-empty diagnostics/errors and equality checks run outside the clock. Ranges give the two trial medians in milliseconds. No editor, cold-start or heap claim.

| Path | N | Before median ms | After median ms |
| --- | ---: | ---: | ---: |
| declared | 1 | 1.0149–1.1430 | 1.0710–1.1219 |
| declared | 100 | 3.6223–3.8044 | 3.5019–3.8687 |
| declared | 1000 | 40.8351–44.1731 | 20.4867–20.9703 |
| held | 1 | 0.8380–0.8954 | 0.8284–0.8752 |
| held | 100 | 3.3918–3.7017 | 3.1899–3.6997 |
| held | 1000 | 38.7595–42.0132 | 18.8838–20.9114 |
| new | 1 | 0.5986–0.7913 | 0.5872–0.6690 |
| new | 100 | 3.9649–4.9670 | 2.7849–2.9273 |
| new | 1000 | 141.5844–145.1629 | 19.8900–20.6235 |

All one-call trial ranges overlap, as do the 100-call declared and held-object ranges; do not infer a consistent improvement there. Membership indexing reads the complete consulted list, whereas an old isolated early match could stop sooner. These fixtures use late matches and do not establish a benefit for isolated early-match queries. Unused membership lists remain lazy.

Reproduce from the repository root:

```powershell
node scripts/benchmark-argument-object-queries.mjs --baseline=ba3c4e5228c4c1a0b25c8b42dab8ba57b49ae37e
node scripts/benchmark-argument-object-queries.mjs
```
