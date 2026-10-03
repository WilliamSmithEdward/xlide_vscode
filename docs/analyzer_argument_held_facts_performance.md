# Lazy held-object facts for argument validation

checkArgumentTypes previously ran heldObjectsAt while creating every procedure visitor. Procedures with no calls or only literal arguments never used those facts. Other consumers already defer held-object setup until needed.

Create the held-object lookup on the first heldClassOf query and retain it in that procedure visitor. Named object and Variant arguments still receive the complete procedure's facts, including preceding assignments and active branches. The cache does not cross visitor or conditional-activity lifetimes.

## Reproduction

```powershell
node scripts/benchmark-argument-held-facts.mjs --baseline=1878e0c7 --rounds=15
node scripts/benchmark-argument-held-facts.mjs --rounds=15
```

To isolate this change with snapshot reuse from PR #804 applied equally to both versions, add `--held-snapshots=8256c3e9` to both commands. That option reads the heldObjects.ts implementation from the supplied Git commit; fetch the PR head if the commit is unavailable locally.

The benchmark times the public argument-type visitor factory and statement traversal after parsing/binding, with warmed caches, three warmups and 15 samples. The no-call fixture declares 100 auto-instantiated Collections and runs ordinary assignments. The held-argument control also passes one named object to a Collection parameter, requiring held facts. The empty control declares no objects. All fixtures require no diagnostics.

Node 24.18.0, AMD Ryzen 7 9800X3D; median milliseconds against main 1878e0c7 (before PR #804):

| Statements | No calls before | After | Held argument before | After | Empty control before | After |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 100 | 0.934 | 0.098 | 0.820 | 0.922 | 0.092 | 0.065 |
| 1,000 | 7.656 | 0.442 | 7.469 | 7.786 | 0.447 | 0.343 |
| 3,000 | 27.166 | 0.999 | 26.636 | 26.355 | 1.272 | 0.980 |

With snapshot reuse applied equally to baseline and optimized argument visitors:

| Statements | No calls before | After | Held argument before | After | Empty control before | After |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 100 | 0.154 | 0.111 | 0.157 | 0.230 | 0.107 | 0.086 |
| 1,000 | 0.667 | 0.446 | 0.730 | 0.802 | 0.449 | 0.324 |
| 3,000 | 1.733 | 0.957 | 1.501 | 1.422 | 1.360 | 0.917 |

These isolated-rule stress timings do not represent editor latency. Absolute values vary with load, and small controls can move between samples. PR #804 reduces the cost of walks that remain necessary; this change avoids unused walks entirely. Warm literal-value/type caches remain outside this change's scope.

## Validation

- Type checking passed; full suite: 601 files, 12,551 tests passed, 13 skipped.
- 83 targeted tests passed.
- 1,500 complete analyzer diagnostic outputs matched baseline 1878e0c7 across parameter types, held values, literals, call forms, branches and conditional compilation.
- Five regressions require zero held-object walks for no calls/literal-only arguments, one walk for repeated named arguments, preserved Variant scalar mismatches and independent activity environments on reused parsed nodes.
