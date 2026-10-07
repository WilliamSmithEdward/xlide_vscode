# Private significant-token filters

Baseline: `bff3b8be7b65042c77dcea490d804ecada8dc85c` (published PR #1210 merge). Measurements used source head `bda73fc4ab0f05e9566bf18b7f224ac2874f4a80`. The rebased analyzer bundle is byte-identical (SHA-256 `09eb2c633602429e6583dd9e4e3fec68a11bc51e4c8a5e11151113f49e7bd796`). Concurrent #1208 changed editor/native recovery files, so types and the full suite were rerun on the combined tree.

Remove four redundant private comment filters across conditionValue and procedureLabels. The condition parser is constructed only at public entries that already remove comments, or from groups of those same tokens. Its builtin groups and argumentValue therefore need no second filter; the redundant outer mapping over groups also disappears. Public conditionValue/numberValue filtering remains necessary and unchanged. Label-reference and completion-prefix helpers receive significant tokens from statementTokens/statementTokensCached. Keep tokensWithoutLeadingLineNumber and its mutable copy contract unchanged. All four private consumers only read the borrowed inputs. No token mutation, new cache or lifetime; returned label metadata retains independent objects/spans.

## Validation

Types and 38 focused tests across four files pass. The 21 new cases comprise nine public work bounds and twelve semantic controls. Baseline fails all nine work bounds and passes all twelve controls. Cases independently assert numeric values, conditions, exact label keys/text/spans/kinds, completion items, frozen input arrays, continuations, numeric labels and trailing comments. Preserve the existing conservative unknown result for two-argument Round; one-argument Round retains banker rounding. Existing depth and label-fact isolation suites pass.

Final full suite on the combined tree: 15,873 passed, 33 skipped across 804 passing files and seven skipped files, in 89.68 seconds. The original isolated run passed 15,866 tests in 98.19 seconds. Complete outputs match 20,000 generated runtime-member public checks, 10,000 generated bundles of numberValue/conditionValue/label-reference/completion queries, 2,000 metadata modules, 10,000 call modules, 2,000 fragment-consumer modules and 24,768 corpus runs over 8,256 sources with LF/CRLF/CR; zero differences. Full comparisons include every diagnostic and internal error.

## Work

Temporarily wrap Array.filter outside timing. Attribute comment predicates to private ConditionParser, labelReferenceGroup or isLabelTargetPrefix frames, excluding the required statementTokens lexer filter. Count elements visited by those redundant predicates and restore the original method in finally before timing. Independent public results are asserted. Completion counts follow the public API's selected prefix bounds: the 1,000-value fixture measures 200 predicate visits, rather than a scan of its whole logical input. Counts describe these public paths, not heap bytes; group outer-map allocations are not counted as predicate visits.

| Fixture | N | Before redundant elements | After |
| --- | ---: | ---: | ---: |
| numeric-builtins | 1 | 2 | 0 |
| label-references | 1 | 1 | 0 |
| completion-prefix | 1 | 3 | 0 |
| numeric-builtins | 100 | 200 | 0 |
| label-references | 100 | 100 | 0 |
| completion-prefix | 100 | 201 | 0 |
| numeric-builtins | 1000 | 2000 | 0 |
| label-references | 1000 | 1000 | 0 |
| completion-prefix | 1000 | 200 | 0 |

## Timings

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor. Sequential baseline/current/current/baseline trials, three warmups and nine measured rounds; ranges are the two trial medians in milliseconds. Fixtures use at most 100 values per physical line with continuation syntax. Public numeric tokenization is excluded, label token lookup/completion parsing included; complete-module setup is included. Exact public results and empty complete diagnostics/internal errors are asserted outside the clock. All four trials have matching full-output hashes. Small/full timing results can be noisy, overlapping or slower. No heap-byte, editor-latency, cold-start or universal speedup claim.

| Fixture | N | Scope | Before median ms | After median ms |
| --- | ---: | --- | ---: | ---: |
| numeric-builtins | 1 | public-query | 0.0069–0.0103 | 0.0087–0.0127 |
| label-references | 1 | public-query | 0.0036–0.0036 | 0.0034–0.0054 |
| completion-prefix | 1 | public-query | 0.0086–0.0096 | 0.0089–0.0089 |
| condition-module | 1 | complete-module-diagnostics | 1.4069–1.4439 | 1.1717–1.2353 |
| numeric-builtins | 100 | public-query | 0.1569–0.1683 | 0.1029–0.1249 |
| label-references | 100 | public-query | 0.0360–0.0483 | 0.0250–0.0389 |
| completion-prefix | 100 | public-query | 0.0886–0.1011 | 0.0669–0.0722 |
| condition-module | 100 | complete-module-diagnostics | 2.2174–2.6496 | 2.1548–2.4643 |
| numeric-builtins | 1000 | public-query | 0.5772–0.7818 | 0.5653–0.5764 |
| label-references | 1000 | public-query | 0.1128–0.1551 | 0.0851–0.1263 |
| completion-prefix | 1000 | public-query | 0.0350–0.0645 | 0.0220–0.0491 |
| condition-module | 1000 | complete-module-diagnostics | 10.1996–10.2375 | 8.4569–10.0092 |

## Reproduction

`node scripts/benchmark-private-significant-token-filters.mjs --baseline=bff3b8be7b65042c77dcea490d804ecada8dc85c`

`node scripts/benchmark-private-significant-token-filters.mjs`

Run sequential ABBA trials. Baseline builds restore only the two changed source files; the surrounding measured analyzer bundle inputs are identical. Inspect independent assertions and matching complete-output hashes before interpreting timings.
