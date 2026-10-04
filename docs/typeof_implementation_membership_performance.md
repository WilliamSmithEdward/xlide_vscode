# TypeOf implementation membership performance

Baseline: `fb60195bc9d6810e093f1ee8ee1a02d5795f85b6` (PR #1178). The measured pre-squash head `41f642c12d84ba3bc687a96a210e64eed640fdc9` has an identical Git tree, verified with `git diff --exit-code` after rebasing this fix onto main.

The TypeOf rule already shared type resolution and interface-sharing queries. Its two compatibility directions still called the unindexed direct-implementation test, repeatedly traversing implemented-name strings. This change passes one lazy query-owned membership lookup to both directions. It shares the existing alias logic and preserves the concrete-operand exclusion, generic/scalar guards, ambiguity and freshness between queries. No new dependency is added.

## Validation

Type checking and 101 focused tests pass. Twelve new bounded-work regressions fail on the baseline, with all 20 prior query controls passing there. Full suite: 15,454 tests passed, 33 skipped, across 788 passing files and seven skipped files, in 74.09 seconds. The rebase replaces only the tree-identical baseline commit.

Complete before/after comparisons: 20,000 generated public-rule outputs, 2,000 generated complete-module outputs, and 24,768 complete-module diagnostic/error results over 8,256 oracle sources with LF/CRLF/CR; zero differences. Generated metadata varies duplicate names, surface kinds, Implemented-name aliases, host/project/generic/scalar/unknown types and case. No diagnostic changes are intended.

## Work counts

Each fixture has N expressions and one N-entry implemented-name list. Repeated/distinct missing project targets independently produce N exact always-False messages/spans. Host-direct and host-reverse fixtures put Worksheet at the end of the project list and independently produce no diagnostics. Proxy array-index reads are counted only outside timing; all timed metadata consists of plain objects and arrays.

| Fixture at N = 1,000 | Before string reads | After string reads |
| --- | ---: | ---: |
| repeated-missing | 2,002,000 | 3,000 |
| distinct-missing | 2,002,000 | 3,000 |
| host-direct | 1,001,000 | 2,000 |
| host-reverse | 1,000,000 | 1,000 |

The remaining linear reads build the distinct membership, interface-sharing and concrete-operand exclusion indexes as needed. These indexes serve different eligibility rules. Unused membership lists remain lazy.

## Timings

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor           . Sequential baseline/current/current/baseline runs, three warmups and nine measured rounds each. The table gives the range of two trial medians in milliseconds. Direct rule setup parses/builds symbols outside its clock; complete-module calls include analyzer setup and capture internal errors. Independent exact diagnostic assertions and complete-output equality run outside the clock. No editor, cold-start or heap improvement is claimed.

| Fixture | N | Scope | Before median ms | After median ms |
| --- | ---: | --- | ---: | ---: |
| repeated-missing | 1 | rule | 0.0109–0.0135 | 0.0141–0.0147 |
| repeated-missing | 1 | complete-module-diagnostics | 1.1485–1.3713 | 1.1658–1.1822 |
| repeated-missing | 100 | rule | 1.4069–1.4794 | 0.1703–0.1790 |
| repeated-missing | 100 | complete-module-diagnostics | 7.9843–9.4115 | 6.7303–6.9623 |
| repeated-missing | 1000 | rule | 133.4387–140.7374 | 1.3472–1.3645 |
| repeated-missing | 1000 | complete-module-diagnostics | 177.9302–182.0192 | 43.7835–47.7716 |
| distinct-missing | 1 | rule | 0.0024–0.0026 | 0.0026–0.0026 |
| distinct-missing | 1 | complete-module-diagnostics | 0.5106–0.5270 | 0.5185–0.6187 |
| distinct-missing | 100 | rule | 1.4422–1.4683 | 0.1910–0.2086 |
| distinct-missing | 100 | complete-module-diagnostics | 5.9465–6.1324 | 4.4467–4.6069 |
| distinct-missing | 1000 | rule | 135.9595–138.7873 | 1.9723–2.1511 |
| distinct-missing | 1000 | complete-module-diagnostics | 183.2840–184.2852 | 46.8071–47.2121 |
| host-direct | 1 | rule | 0.0021–0.0022 | 0.0024–0.0027 |
| host-direct | 1 | complete-module-diagnostics | 0.5131–0.5211 | 0.4947–0.5001 |
| host-direct | 100 | rule | 0.6911–0.7320 | 0.0915–0.1037 |
| host-direct | 100 | complete-module-diagnostics | 4.9551–4.9680 | 4.2756–4.4402 |
| host-direct | 1000 | rule | 66.9037–70.0256 | 0.7547–0.7734 |
| host-direct | 1000 | complete-module-diagnostics | 109.2029–111.8128 | 40.4280–40.6753 |
| host-reverse | 1 | rule | 0.0014–0.0015 | 0.0016–0.0017 |
| host-reverse | 1 | complete-module-diagnostics | 0.4917–0.5115 | 0.4865–0.5170 |
| host-reverse | 100 | rule | 0.6810–0.7173 | 0.0641–0.0643 |
| host-reverse | 100 | complete-module-diagnostics | 4.9205–4.9354 | 4.1780–4.2058 |
| host-reverse | 1000 | rule | 67.5175–70.6295 | 0.5987–0.6051 |
| host-reverse | 1000 | complete-module-diagnostics | 108.6116–111.7386 | 39.8376–41.0732 |

Small cases are mixed: repeated-missing direct calls are slightly slower; their full-module trial ranges overlap. Distinct-missing one-expression full calls are slower, and the direct range touches the baseline. Host-direct and host-reverse one-expression direct calls are also slightly slower; host-reverse full ranges overlap. Building the membership set visits a whole consulted list, whereas an old early match could stop before its end. The benchmark uses missing or late-match lists and does not establish a benefit for an isolated early-match query.

Reproduce from the repository root:

```powershell
node scripts/benchmark-typeof-implementation-membership.mjs --baseline=fb60195bc9d6810e093f1ee8ee1a02d5795f85b6
node scripts/benchmark-typeof-implementation-membership.mjs
```
