# Runtime member metadata query performance

Baseline: `1536ea7fcc401e675980749d99b752f726210cb9` (published PR #1201 merge). Measurements used source head `e91bc4c12fc3b6bbace6bbc0cbdef494290d8537`; its entire tracked baseline tree was verified identical to the published merge. Baseline builds restore only rules/lateBoundMembers.ts; surrounding files are identical.

WorksheetFunction name sets are rebuilt per matching reference; ActiveSheet's combined host/document name set is rebuilt per statement; form control filtering and name search are repeated per Controls reference. Share lazy worksheet-function/sheet queries and per-form control summaries within each public rule invocation. Control counts retain duplicates; control-name sets are built only when a literal-name query needs one. Preserve host/model/completeness/shadowing/added-control guards. Fresh invocations rebuild queries, even with retained AST/metadata. No cross-invocation cache is introduced. Duplicate token allocation is a separate, unmeasured candidate.

## Validation

Types and 46 focused tests pass. Thirty-six new cases cover 18 public/full work bounds at 10/100/1,000 metadata entries and references; three invocation-wide bounds across 100 procedures; 15 semantic controls covering known names/casing, Excel/host-completeness gates, document-only sheet authority, shadowed ActiveSheet, control filtering/casing, duplicate numeric counts, incomplete/dynamic-control guards and retained-AST model/form freshness. Baseline fails all 21 work bounds and passes all 15 semantic controls. Every work case compares complete independently expected code/message/spans; full cases also require empty internal-error lists.

Full suite: 15,802 passed, 33 skipped across 800 passing files and seven skipped files, in 80.25 seconds.

Complete outputs match 20,000 generated public-rule results, 2,000 generated modules and 24,768 corpus runs over 8,256 sources with LF/CRLF/CR; zero differences. Generated cases vary host/application gates, known/unknown/hidden function names, sheet host/document presence, form kinds/completeness/duplicate return types, literal/numeric/dynamic controls, shadowing, cached/uncached contexts, auto-instantiation, multiple procedures, conditional/block/With forms. Full comparisons include all diagnostics and internal errors.

## Work

Fixtures have N metadata entries and N missing-name references. Form receivers are As New Form1 to avoid separate Nothing diagnostics. Host name and form return-type getters run only outside timing. Getter properties become plain properties before timing. Expected output is exactly N runtime-member-not-found diagnostics, with exact text/spans and no full-module internal errors.

| N = 1,000 | API | Before name reads | After name reads | Before return reads | After return reads |
| --- | --- | ---: | ---: | ---: | ---: |
| worksheetFunction | public-rule | 1,002,000 | 3,000 | 0 | 0 |
| worksheetFunction | complete-module-diagnostics | 1,002,000 | 3,000 | 0 | 0 |
| activeSheet | public-rule | 1,001,000 | 2,000 | 0 | 0 |
| activeSheet | complete-module-diagnostics | 1,002,000 | 3,000 | 0 | 0 |
| formControls | public-rule | 0 | 0 | 1,000,000 | 1,000 |
| formControls | complete-module-diagnostics | 0 | 0 | 1,000,000 | 1,000 |

This bounds the repaired metadata queries in these fixtures, not every analyzer path.

## Timings

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor. Sequential baseline/current/current/baseline trials; three warmups and nine measured rounds. Ranges are the two trial medians in milliseconds. Host model indexes are warm during timing. Public parsing/symbol construction is excluded; full analyzer setup and fresh query construction are included. Assertions run outside the clock. No editor/heap/cold-start claim.

| Metadata | API | N | Before median ms | After median ms |
| --- | --- | ---: | ---: | ---: |
| worksheetFunction | public-rule | 1 | 0.0208–0.0241 | 0.0173–0.0188 |
| worksheetFunction | complete-module-diagnostics | 1 | 0.8843–0.9215 | 0.9670–0.9698 |
| worksheetFunction | public-rule | 100 | 1.3249–1.3692 | 0.8037–0.8102 |
| worksheetFunction | complete-module-diagnostics | 100 | 4.6002–4.8464 | 4.5285–4.6182 |
| worksheetFunction | public-rule | 1000 | 40.8341–41.9181 | 7.4154–7.4885 |
| worksheetFunction | complete-module-diagnostics | 1000 | 62.5683–63.2631 | 28.8859–29.5501 |
| activeSheet | public-rule | 1 | 0.0029–0.0049 | 0.0031–0.0031 |
| activeSheet | complete-module-diagnostics | 1 | 0.3845–0.3953 | 0.3897–0.3933 |
| activeSheet | public-rule | 100 | 0.4458–0.4466 | 0.0700–0.0702 |
| activeSheet | complete-module-diagnostics | 100 | 3.1053–3.1815 | 2.9226–2.9955 |
| activeSheet | public-rule | 1000 | 33.8508–34.9980 | 0.6250–0.6428 |
| activeSheet | complete-module-diagnostics | 1000 | 60.5542–62.5664 | 24.9708–26.2909 |
| formControls | public-rule | 1 | 0.0052–0.0052 | 0.0056–0.0057 |
| formControls | complete-module-diagnostics | 1 | 0.7018–0.7105 | 0.6975–0.7006 |
| formControls | public-rule | 100 | 1.2329–1.2813 | 0.7760–0.8156 |
| formControls | complete-module-diagnostics | 100 | 5.2123–5.2432 | 4.6250–4.6718 |
| formControls | public-rule | 1000 | 37.2394–37.7138 | 6.9313–6.9973 |
| formControls | complete-module-diagnostics | 1000 | 64.5634–66.3526 | 36.5197–37.1961 |

All three 1,000-reference fixtures improve in both public and full-module timing trials. Single-reference WorksheetFunction full diagnostics and form-control public checks are slightly slower in both trials; single ActiveSheet ranges overlap, as do 100-reference WorksheetFunction complete-module ranges. These results do not establish an improvement for every small module.

Reproduce sequentially from the repository root:

```powershell
node scripts/benchmark-runtime-member-metadata.mjs --baseline=1536ea7fcc401e675980749d99b752f726210cb9
node scripts/benchmark-runtime-member-metadata.mjs
node scripts/benchmark-runtime-member-metadata.mjs
node scripts/benchmark-runtime-member-metadata.mjs --baseline=1536ea7fcc401e675980749d99b752f726210cb9
```
