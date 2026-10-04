# Late-bound argument comment filters

Baseline: `f70281edcb9ebdb779b680da3749cfd75e28de99` (published PR #1205 merge). Measurements used `2b4e5910cf40a416b132e3a5f555e3aa94a65664`; its runtime, tests and benchmark scripts are identical. Only the readonly-copy report baseline pin changed after rebasing onto published #1203.

The private argumentRefusal helper only receives statementTokensAfterLeadingLabel tokens. Its statementTokensCached input has already removed comments and newlines. Splitting that input preserves the significant-token contract. Remove two whole-input comment filters in bare calls and one per-argument filter in the required-parameter check. The named-argument filter still performs necessary work and remains unchanged. Token identities, groups, spans and diagnostics are preserved; no new cache or lifetime is introduced.

## Validation

Types and 59 focused tests across four files pass. The 17 new tests comprise six public work bounds and 11 semantic controls covering comments, continuations, labels, colons, nested/parenthesized arguments, named arguments, omitted arguments and required parameters. Baseline fails all six work bounds and passes all 11 semantic controls. Work assertions independently check exact public diagnostics, including the runtime 450 message/span and quiet nested-argument controls.

Full suite: 15,824 passed, 33 skipped across 802 passing files and seven skipped files, in 93.03 seconds. Complete outputs match 20,000 generated public checks, 2,000 generated metadata modules, 10,000 generated call modules and 24,768 corpus runs over 8,256 sources using LF/CRLF/CR; zero differences. Comparisons include every diagnostic and internal error.

## Work

Counters attribute token-kind reads to Array.filter in argumentRefusal using temporary getters only outside timing. Restore descriptors in finally before measuring normal token properties. These counts describe this public path, not heap bytes or editor latency. Both whole-input passes allocate temporary arrays; the required-parameter filter also allocates an array for each group inspected.

| Fixture | N | Source tokens | Before filter kind reads | After |
| --- | ---: | ---: | ---: | ---: |
| bare-values | 1 | 4 | 9 | 0 |
| bare-values | 100 | 202 | 404 | 0 |
| bare-values | 1000 | 2002 | 4004 | 0 |
| nested-one-argument | 1 | 7 | 18 | 0 |
| nested-one-argument | 100 | 205 | 612 | 0 |
| nested-one-argument | 1000 | 2005 | 6012 | 0 |

## Timings

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor. Sequential baseline/current/current/baseline trials, three warmups and nine measured rounds. Ranges are two trial medians in milliseconds. Public parsing/symbol setup is excluded; complete-module setup is included. Exact public/primary full findings, complete snapshots and internal errors are asserted outside the clock. All four trials have matching full-output hashes. These small isolated timings are mixed and do not establish a universal speedup, heap-byte saving, cold-start improvement or editor-latency improvement.

| Fixture | Scope | N | Before median ms | After median ms |
| --- | --- | ---: | ---: | ---: |
| bare-values | public-rule | 1 | 0.0982–0.1357 | 0.1095–0.1109 |
| bare-values | complete-module-diagnostics | 1 | 1.5172–1.6436 | 1.2915–1.3886 |
| bare-values | public-rule | 100 | 0.0982–0.1024 | 0.0835–0.0843 |
| bare-values | complete-module-diagnostics | 100 | 1.8164–1.9489 | 1.7584–1.8366 |
| bare-values | public-rule | 1000 | 0.3527–0.5410 | 0.3102–0.3402 |
| bare-values | complete-module-diagnostics | 1000 | 6.2833–6.4030 | 6.3690–6.7429 |
| nested-one-argument | public-rule | 1 | 0.0507–0.0514 | 0.0501–0.0502 |
| nested-one-argument | complete-module-diagnostics | 1 | 1.0605–1.0918 | 1.0306–1.0685 |
| nested-one-argument | public-rule | 100 | 0.0972–0.1119 | 0.0828–0.0852 |
| nested-one-argument | complete-module-diagnostics | 100 | 1.6979–1.9060 | 1.9595–1.9665 |
| nested-one-argument | public-rule | 1000 | 0.3296–0.6216 | 0.2511–0.2695 |
| nested-one-argument | complete-module-diagnostics | 1000 | 8.2489–8.5559 | 7.5894–7.7710 |

## Reproduction

`node scripts/benchmark-late-bound-comment-filters.mjs --baseline=f70281edcb9ebdb779b680da3749cfd75e28de99`

`node scripts/benchmark-late-bound-comment-filters.mjs`

Run sequential ABBA trials. The baseline loader restores only lateBoundMembers.ts; surrounding runtime sources are identical. Inspect independent output assertions and matching complete-output hashes before interpreting timings.
