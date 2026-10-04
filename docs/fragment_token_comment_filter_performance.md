# Fragment-token comment filters

Baseline: `04f3e50d03839211f945e26e986d075f78eeda22` (published PR #1207 merge). Measurements used source head `8a90ff30ec633251b011dfae9053776b497193db`; its entire tracked tree matches the published merge.

rawExpressionTokens delegates to statementTokens, which already removes comments and newlines. The leading-date-literal path removes its synthetic equals token and adjusts offsets without reintroducing trivia. Remove eleven redundant comment filters across eight analyzer files: typeInference (2), arrays, collectionState, lateBoundMembers, omittedArguments, overflow (2), shared and typeFieldArrays (2). These filters each scan significant tokens and create another array without changing its contents. Preserve the canonical lexer filtering, token identities, spans, mutable fragment arrays and existing public behavior; no new cache or lifetime.

## Validation

Types and 82 focused tests across six files pass. The 21 new tests include six public work bounds and 15 independent semantic/contract controls. Baseline fails all six work bounds and passes all 15 controls. Controls include constants, source-expression syntax, continuations, dates through the prefixed lexer path, apostrophes in strings, trailing comments and newlines.

Full suite: 15,845 passed, 33 skipped across 803 passing files and seven skipped files, in 87.32 seconds. Complete outputs match 20,000 generated public checks, 2,000 metadata modules, 10,000 call modules, 2,000 additional fragment-consumer modules and 24,768 corpus runs over 8,256 sources with LF/CRLF/CR; zero differences. Fragment modules cover string/numeric constants, optional defaults, array bounds, UDT field arrays, host For Each and date constants/defaults. Comparisons include every diagnostic and internal error.

## Redundant work

Temporarily wrap Array.filter only outside timing. Attribute comment predicates to the selected public query and exclude the required statementTokens lexer filter using its stack. Count array elements visited by those redundant predicates; restore the original method in finally before timing. Independently assert source-expression syntax and constant-string results. Counts describe two public paths, not heap bytes or all eleven removed sites.

| Fixture | N | Before redundant elements | After |
| --- | ---: | ---: | ---: |
| source-expression | 1 | 4 | 0 |
| constant-string | 1 | 1 | 0 |
| source-expression | 100 | 202 | 0 |
| constant-string | 100 | 199 | 0 |
| source-expression | 1000 | 2002 | 0 |
| constant-string | 1000 | 1999 | 0 |

## Timings

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor. Sequential baseline/current/current/baseline trials, three warmups and nine measured rounds; ranges are the two trial medians in milliseconds. Every fixture has at most 100 values per physical line with valid continuation syntax. Public symbol setup is excluded, fragment lexing included; complete-module setup is included. Independent query results and empty complete diagnostics/internal errors are asserted outside the clock; all four trials match full-output hashes. Small timings are noisy and mixed. No heap-byte, editor-latency, cold-start or universal speedup claim.

| Fixture | N | Scope | Before median ms | After median ms |
| --- | ---: | --- | ---: | ---: |
| source-expression | 1 | public-query | 0.0045–0.0048 | 0.0045–0.0072 |
| constant-string | 1 | public-query | 0.0014–0.0016 | 0.0013–0.0022 |
| constant-module | 1 | complete-module-diagnostics | 1.0553–1.0781 | 1.1052–1.1867 |
| source-expression | 100 | public-query | 0.0620–0.0623 | 0.0538–0.0714 |
| constant-string | 100 | public-query | 0.0359–0.0377 | 0.0376–0.0389 |
| constant-module | 100 | complete-module-diagnostics | 1.1916–1.2515 | 1.2375–1.2397 |
| source-expression | 1000 | public-query | 0.2488–0.2541 | 0.2277–0.2279 |
| constant-string | 1000 | public-query | 0.2114–0.2165 | 0.1867–0.2113 |
| constant-module | 1000 | complete-module-diagnostics | 3.6370–4.6598 | 3.9777–4.5665 |

## Reproduction

`node scripts/benchmark-fragment-token-comment-filters.mjs --baseline=04f3e50d03839211f945e26e986d075f78eeda22`

`node scripts/benchmark-fragment-token-comment-filters.mjs`

Run sequential ABBA trials. Baseline builds restore all eight changed source files; surrounding runtime sources are identical. Inspect independent assertions and complete-output hashes before interpreting timings.
