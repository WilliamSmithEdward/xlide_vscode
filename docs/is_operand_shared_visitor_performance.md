# Shared Is operand visitor

Baseline: `bfedb82ca333c89591fbaa649ede46ffbe7bca6b` (published PR #1212 merge). Measurements used source head `19da0cc97e17a1e2244a7d71b4cb4a0e098bf5b8`; its entire tracked tree matches the published merge.

The parsed-expression and separately parsed condition paths contained identical Is operand diagnostics. Reuse checkIsOperatorOperands' existing per-procedure visitor in checkIsOperandsInConditions. Keep one factory per invocation, one readonly type environment/visitor per active procedure, existing expression traversal and exact diagnostic text/order/spans. Remove the condition path's dead comment filter: statementTokens already excludes comments/newlines. Its required absolute-offset token mapping remains unchanged. No new cache, ownership or lifetime.

A normalized AST scan of all 193 analyzer files found this single exact pair among 2,025 functions with at least three top-level statements. After sharing the visitor, the same scan finds no exact pair among 2,024 such functions. This is narrow exact-body evidence; it does not prove the absence of semantic duplication or dead exported code.

## Validation

Types and 72 focused tests across four files pass. Thirteen new tests include three public work tests covering two independent bounds each, plus ten semantic controls. Baseline fails all three work tests on both visitor identity and redundant filtering, and passes all ten controls. At 1,000 scalar conditions, distinct callbacks fall from 1,000 to one and redundant token checks from 9,000 to zero. Independent exact public warnings include text and source spans. Semantic controls cover single-line If, Do While/Until, Loop While/Until, While/Wend, Object/Variant/user-class controls and separate per-procedure environments.

Full suite: 15,886 passed, 33 skipped across 805 passing files and seven skipped files, in 91.25 seconds. Complete outputs match 10,000 generated Is-rule queries and 2,000 corresponding full modules, 20,000 runtime-member queries, 10,000 four-API private query bundles, 2,000 metadata modules, 10,000 call modules, 2,000 fragment modules and 24,768 corpus runs over 8,256 sources with LF/CRLF/CR; zero differences. Full comparisons include every diagnostic and internal error.

## Filter work

Temporarily wrap Array.filter outside timing, attribute comment predicates to the condition checker and exclude the required statementTokens lexer filtering. Restore the original method in finally. Assert independent public warnings/quiet Object results. Callback identity is checked separately by spying on expression traversal in tests; timing uses ordinary uninstrumented traversal. Counts do not measure heap bytes.

| Type | Conditions | Before redundant checks | After |
| --- | ---: | ---: | ---: |
| Long | 0 | 0 | 0 |
| Long | 1 | 9 | 0 |
| Long | 100 | 900 | 0 |
| Long | 1000 | 9000 | 0 |
| Object | 0 | 0 | 0 |
| Object | 1 | 9 | 0 |
| Object | 100 | 900 | 0 |
| Object | 1000 | 9000 | 0 |

## Timings

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor. Sequential baseline/current/current/baseline trials, three warmups and nine measured rounds; ranges are two trial medians in milliseconds. Include zero-condition setup controls as well as scalar-warning and quiet Object paths. Public parsing/symbol setup is excluded; complete-module setup is included. Independent exact public/primary full diagnostics and complete snapshots/internal errors are asserted outside timing; all four trials match full-output hashes. Small/empty/full results can overlap or regress. No universal speedup, heap-byte, editor-latency or cold-start claim.

| Type | Conditions | Scope | Before median ms | After median ms |
| --- | ---: | --- | ---: | ---: |
| Long | 0 | public-rule | 0.0012–0.0018 | 0.0014–0.0015 |
| Long | 0 | complete-module-diagnostics | 0.7191–0.7368 | 0.7083–0.7130 |
| Long | 1 | public-rule | 0.0078–0.0082 | 0.0075–0.0077 |
| Long | 1 | complete-module-diagnostics | 0.9732–0.9971 | 0.9511–1.0023 |
| Long | 100 | public-rule | 0.1632–0.1660 | 0.1462–0.2096 |
| Long | 100 | complete-module-diagnostics | 6.5056–6.7866 | 6.3985–7.0100 |
| Long | 1000 | public-rule | 0.7033–0.7060 | 0.6643–0.8961 |
| Long | 1000 | complete-module-diagnostics | 40.8030–41.0677 | 41.7260–42.1279 |
| Object | 0 | public-rule | 0.0003–0.0003 | 0.0005–0.0005 |
| Object | 0 | complete-module-diagnostics | 0.3748–0.3837 | 0.3862–0.3973 |
| Object | 1 | public-rule | 0.0011–0.0011 | 0.0010–0.0010 |
| Object | 1 | complete-module-diagnostics | 0.6053–0.6177 | 0.6131–0.6279 |
| Object | 100 | public-rule | 0.0564–0.0566 | 0.0456–0.0457 |
| Object | 100 | complete-module-diagnostics | 5.3092–5.4024 | 5.1481–5.2753 |
| Object | 1000 | public-rule | 0.5071–0.5282 | 0.4191–0.4191 |
| Object | 1000 | complete-module-diagnostics | 43.7290–44.4825 | 42.8499–43.9243 |

## Reproduction

`node scripts/benchmark-is-operand-shared-visitor.mjs --baseline=bfedb82ca333c89591fbaa649ede46ffbe7bca6b`

`node scripts/benchmark-is-operand-shared-visitor.mjs`

Run sequential ABBA trials. Baseline builds restore only typeOfIs.ts; surrounding analyzer sources are identical. Inspect output assertions and matching complete-output hashes before interpreting timings.
