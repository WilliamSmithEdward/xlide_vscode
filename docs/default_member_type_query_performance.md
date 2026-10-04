# Default-member type query performance

Baseline: `7feca913b252eefb0aa4ef9f19863c7213ee246c` (PR #1184). All three production files are restored for baseline comparisons. This baseline precedes the separate module parameter classification fix in PR #1187; those surrounding facts are identical in both benchmark sides.

Default-value and statement-form rules repeatedly resolve project types and their default-member verdicts across procedures. Each public rule now owns a lazy object-type/default-verdict/index-requirement query. Object-value paths reuse it for declared objects, held values, created objects and host chains. Project lookups resume from their last position and cache encountered names, including misses and incomplete first classes. The class-only absence check retains its own first-class semantics; default metadata retains the existing first-named surface of any kind. Optional helper parameters keep the original unindexed path for callers that omit them. Facts belong to one invocation; later invocations see refreshed metadata. No dependency is added.

## Validation

Types and 67 focused tests pass. Twenty-nine new tests include 12 public-rule work bounds at 10/100/1,000 procedures with repeated or distinct types, eight verdict/index controls, six first-surface-kind controls, an incomplete-first-class public-rule control, metadata freshness and lazy unused member arrays. On the baseline, eight work bounds fail, four small work bounds and the incomplete-first public control pass; 16 factory-helper tests are intentionally skipped because the new API does not exist there.

Full suite: 15,521 passed, 33 skipped across 791 passing files and seven skipped files, in 88.77 seconds. Complete outputs match 20,000 generated public-rule result bundles, 2,000 generated full-module results and 24,768 corpus diagnostic/internal-error comparisons over 8,256 sources with LF/CRLF/CR; zero differences. Generated cases vary metadata kinds, duplicates, completeness, default signatures, required/optional/index arguments, scalar/project/host/unknown types, plain/indexed/coerced/conditional/New/held/concatenated reads. No diagnostic changes are intended.

## Work

Each fixture has N procedures with an object parameter, N+1 class surfaces and one Debug.Print read per procedure. Expected results independently assert N exact default-value messages and spans, no statement-form findings and no full-module internal errors. Getter counters are untimed.

| N = 1,000 types | Scope | Before class-name reads | After class-name reads |
| --- | --- | ---: | ---: |
| repeated | values-rule | 1,003,001 | 1,004 |
| repeated | statements-rule | 1,003,000 | 1,003 |
| repeated | complete-module-diagnostics | 3,013,006 | 1,009,012 |
| distinct | values-rule | 2,003,000 | 4,001 |
| distinct | statements-rule | 1,502,500 | 3,001 |
| distinct | complete-module-diagnostics | 4,512,505 | 1,014,007 |

Both public rules have linear name-read bounds in these fixtures. Complete-module work retains parameter classification from the pinned baseline; PR #1187 addresses that separately. These measurements do not establish a linear bound for every full-analyzer consumer.

## Timings

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor. Sequential baseline/current/current/baseline trials, three warmups and nine measured rounds. Ranges are the two trial medians in milliseconds. Public-rule timings exclude parsing/symbol construction; complete-module timings include analyzer setup and internal-error capture. Metadata getters are replaced with plain properties before timing. Independent output expectations and equality checks run outside the clock. No editor, cold-start or heap claim.

| Types | Scope | N | Before median ms | After median ms |
| --- | --- | ---: | ---: | ---: |
| repeated | values-rule | 1 | 0.0281–0.0296 | 0.0251–0.0257 |
| repeated | statements-rule | 1 | 0.0108–0.0176 | 0.0114–0.0119 |
| repeated | complete-module-diagnostics | 1 | 1.0998–1.1126 | 1.0231–1.3344 |
| repeated | values-rule | 100 | 0.5507–0.5571 | 0.3164–0.3233 |
| repeated | statements-rule | 100 | 0.4079–0.4212 | 0.1729–0.1974 |
| repeated | complete-module-diagnostics | 100 | 8.1705–9.3689 | 8.6200–10.2043 |
| repeated | values-rule | 1000 | 23.5166–25.4571 | 7.0123–7.0661 |
| repeated | statements-rule | 1000 | 20.0023–21.9100 | 3.8567–4.2373 |
| repeated | complete-module-diagnostics | 1000 | 136.2889–154.6168 | 106.5025–111.6568 |
| distinct | values-rule | 1 | 0.0031–0.0047 | 0.0030–0.0031 |
| distinct | statements-rule | 1 | 0.0018–0.0021 | 0.0022–0.0023 |
| distinct | complete-module-diagnostics | 1 | 0.3436–0.3526 | 0.3487–0.3628 |
| distinct | values-rule | 100 | 0.8314–1.0643 | 0.3623–0.3756 |
| distinct | statements-rule | 100 | 0.4550–0.4916 | 0.2209–0.2311 |
| distinct | complete-module-diagnostics | 100 | 8.4253–8.4767 | 5.4009–5.9626 |
| distinct | values-rule | 1000 | 37.8199–41.3987 | 7.7065–11.5860 |
| distinct | statements-rule | 1000 | 28.1095–29.0663 | 4.9940–6.1883 |
| distinct | complete-module-diagnostics | 1000 | 164.0793–176.3042 | 105.3900–113.8022 |

The one-procedure complete-module ranges overlap. The repeated 100-procedure full-module trials are mixed, including a slower after trial; distinct one-procedure statement-form trials are also slower. These results do not establish a universal small-case gain. The large public-rule and complete-module fixtures improve in both trials.

Reproduce from the repository root, sequentially:

```powershell
node scripts/benchmark-default-member-type-queries.mjs --baseline=7feca913b252eefb0aa4ef9f19863c7213ee246c
node scripts/benchmark-default-member-type-queries.mjs
node scripts/benchmark-default-member-type-queries.mjs
node scripts/benchmark-default-member-type-queries.mjs --baseline=7feca913b252eefb0aa4ef9f19863c7213ee246c
```
