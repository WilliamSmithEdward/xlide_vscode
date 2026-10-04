# Raw project member query performance

Baseline: `babc9e26f0335687518b3d9793c3b73274db44c7` (PR #1199 source head). Baseline builds restore only completion/memberAccess.ts; surrounding files are identical.

Raw project-member queries scan from the start for every occurrence. Share a lazy, resumable first-match index across raw member, signature and project/combined receiver-return lookups, scoped to the caller's existing immutable per-pass memberSurfaceCache. Each project's index stops at the first requested match and resumes for new names. Misses exhaust it once. Raw queries do not inherit host/control members, and first case-insensitive duplicates preserve their complete metadata object. Without a supplied pass cache, retain the existing direct find for cheap first hits and mutable metadata freshness. No global metadata lifetime is introduced.

## Validation

Types and 163 focused tests pass. Twenty-six new cases include nine public first/last/distinct work bounds at 10/100/1,000 members; three complete-module write bounds; six raw kind controls; three duplicate/full-metadata controls; resumed-index/miss work; uncached first-hit/mutation; fresh-cache retained metadata; implicit controls excluded; argument-sensitive return chain. Baseline fails nine work checks, passes four small/first-hit work checks and all 13 semantic controls. Public lookups assert exact metadata object identity. Full diagnostics assert exact N Set-required code/message/spans and empty internal-error lists.

Full suite: 15,766 passed, 33 skipped across 799 passing files and seven skipped files, in 79.35 seconds.

Complete outputs match 10,000 generated completion/menu/raw-member/receiver tuples, 20,000 rule bundles, 2,000 complete modules and 24,768 corpus runs over 8,256 sources with LF/CRLF/CR; zero differences. Generated cases retain host/combined/union/runtime and duplicate/hidden/signature/kind/cached/uncached/metadata variations. No diagnostic changes are intended.

## Work

Class1 has N signatureless Object properties. Public raw queries use a fresh pass cache and compare exact selected metadata identity. Full diagnostics analyze N distinct writes and include setup. Getter instrumentation runs only outside timing; plain names replace it before timing. Expected metadata/diagnostic comparisons and empty errors run outside the clock.

| Scope | N | Before member-name reads | After member-name reads |
| --- | ---: | ---: | ---: |
| raw-first-pass | 1 | 1 | 1 |
| raw-last-pass | 1 | 1 | 1 |
| raw-distinct-pass | 1 | 1 | 1 |
| complete-module-diagnostics | 1 | 7 | 7 |
| raw-first-pass | 100 | 100 | 1 |
| raw-last-pass | 100 | 10,000 | 100 |
| raw-distinct-pass | 100 | 5,050 | 100 |
| complete-module-diagnostics | 100 | 5,650 | 700 |
| raw-first-pass | 1000 | 1,000 | 1 |
| raw-last-pass | 1000 | 1,000,000 | 1,000 |
| raw-distinct-pass | 1000 | 500,500 | 1,000 |
| complete-module-diagnostics | 1000 | 506,500 | 7,000 |

The repaired pass-scoped member queries are bounded in these fixtures. Uncached direct queries and unrelated analyzer work remain separate; no universal linear-work claim.

## Timings

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor. Sequential baseline/current/current/baseline trials, three warmups and nine measured rounds. Ranges are the two trial medians in milliseconds. Parsing is excluded from raw-query timings, included in full diagnostics. Fresh per-pass index construction is timed. No editor/heap/cold-start claim.

| Scope | N | Before median ms | After median ms |
| --- | ---: | ---: | ---: |
| raw-first-pass | 1 | 0.0075–0.0093 | 0.0098–0.0109 |
| raw-last-pass | 1 | 0.0019–0.0020 | 0.0021–0.0022 |
| raw-distinct-pass | 1 | 0.0020–0.0020 | 0.0022–0.0026 |
| complete-module-diagnostics | 1 | 1.2067–1.2469 | 1.0811–1.1882 |
| raw-first-pass | 100 | 0.1172–0.1214 | 0.1191–0.1455 |
| raw-last-pass | 100 | 0.2422–0.2584 | 0.1064–0.1564 |
| raw-distinct-pass | 100 | 0.1107–0.1119 | 0.0933–0.1044 |
| complete-module-diagnostics | 100 | 4.1005–4.1054 | 4.3306–4.5114 |
| raw-first-pass | 1000 | 0.5625–0.7199 | 0.7682–1.0591 |
| raw-last-pass | 1000 | 8.5903–9.4510 | 0.5947–0.8155 |
| raw-distinct-pass | 1000 | 4.5578–4.7740 | 0.7298–1.1998 |
| complete-module-diagnostics | 1000 | 30.4817–31.0175 | 24.8063–28.1465 |

At 1,000 members, last/distinct raw queries and complete diagnostics improve in both trials. Repeated first-member raw queries at 1,000 members are slower despite fewer getter reads: the old first-hit find is cheap and the pass index has overhead. All single raw-query cases are slightly slower; 100-write full diagnostics are slower in both trials; 100-first-query ranges overlap. This is a repair for repeated expensive member searches, not evidence that every lookup or module improves.

Reproduce sequentially from the repository root:

```powershell
node scripts/benchmark-project-member-queries.mjs --baseline=babc9e26f0335687518b3d9793c3b73274db44c7
node scripts/benchmark-project-member-queries.mjs
node scripts/benchmark-project-member-queries.mjs
node scripts/benchmark-project-member-queries.mjs --baseline=babc9e26f0335687518b3d9793c3b73274db44c7
```
