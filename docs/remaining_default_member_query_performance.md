# Remaining default-member query performance

Baseline: `edb0d30923fe30ffaf5c3a67dd0b03b02b543b08` (PR #1193). All six production files are restored for baseline comparisons.

Assignment type facts, grouping parentheses and scalar argument reads repeatedly resolve object/default-member metadata. Each public rule now owns one lazy default-query bundle. Assignment facts reuse its type resolver and verdicts; read-only project defaults use its resumable first-class lookup. Assignment Let-state calls pass the same optional bundle through cold object-state walks; callers omitting it retain their existing path. Parentheses and scalar New/ByVal argument reads share cached index requirements. Argument query bundles accept an optional needsIndex helper, preserving legacy bundles and omitted options.

The shared factory owns separate any-kind and class-only lookups, each preserving the first matching surface, including an incomplete first class. Object-value absence checks reuse that class-only lookup instead of maintaining a duplicate scan. No cache lifetime is extended beyond a public invocation, existing walk cache ownership is unchanged, later invocations see refreshed metadata and unused member arrays remain lazy. No dependency is added.

## Validation

Types and 163 focused tests pass. Thirty-nine new tests include 30 public-rule bounds: cold/primed no-default assignments and writable/read-only defaults at 10/100/1,000 distinct types, plus grouping parentheses and scalar variable/New arguments with repeated/distinct types. Work tests assert complete exact Set-required/read-only findings or independently empty results. Eight new factory controls cover first-class/any-kind distinctions, incomplete-first metadata, new-query freshness and lazy unused member arrays; one public assignment control checks retained-AST metadata freshness. The baseline fails 27 work bounds, passes three small work cases and the freshness control; eight new factory API tests are intentionally skipped there.

Full suite: 15,698 passed, 33 skipped across 796 passing files and seven skipped files, in 123.01 seconds. Complete outputs match 20,000 generated public-rule bundles (assignment, parentheses, argument and object-value results), 2,000 generated complete modules and 24,768 corpus diagnostic/internal-error comparisons over 8,256 sources with LF/CRLF/CR; zero differences. Generated cases vary surface kinds/completeness/duplicates, default signatures/writable/accessor flags, local/module/array/function/parameter forms, scalar/generic/project/host/unknown types, Let/grouped/indexed/value/operator/coercion reads, scalar variable/New/Call arguments, guards and block headers. No diagnostic changes are intended.

## Work

Fixtures have N procedures and N+1 class surfaces, with repeated or distinct types. Assignment fixtures use one Let per procedure and vary no/writable/read-only defaults. Parentheses and scalar argument fixtures use one corresponding read per procedure. Primed assignments fill object-state walks on the same symbols/context before measuring; cold assignments use unfilled walks. Counters are untimed. Public results independently assert exact N Set-required/read-only findings or no findings. Complete-module expectations independently assert exact code/message/span tuples for Set-required, unset default assignments, read-only assignments, assigned-but-unread and unused declarations, plus no internal errors.

| N = 1,000 mode | Types | Scope | Before class-name reads | After class-name reads |
| --- | --- | --- | ---: | ---: |
| cold-no-default | repeated | public-rule | 1,005,005 | 1,003 |
| cold-no-default | repeated | complete-module-diagnostics | 8,016 | 7,014 |
| cold-no-default | distinct | public-rule | 4,007,000 | 3,001 |
| cold-no-default | distinct | complete-module-diagnostics | 2,515,506 | 14,007 |
| primed-no-default | repeated | public-rule | 2,005 | 1,003 |
| primed-no-default | repeated | complete-module-diagnostics | 8,016 | 7,014 |
| primed-no-default | distinct | public-rule | 2,504,500 | 3,001 |
| primed-no-default | distinct | complete-module-diagnostics | 2,515,506 | 14,007 |
| writable-default | repeated | public-rule | 2,006 | 1,004 |
| writable-default | repeated | complete-module-diagnostics | 8,016 | 7,014 |
| writable-default | distinct | public-rule | 3,005,000 | 4,001 |
| writable-default | distinct | complete-module-diagnostics | 3,015,006 | 14,007 |
| readonly-default | repeated | public-rule | 2,006 | 1,004 |
| readonly-default | repeated | complete-module-diagnostics | 8,016 | 7,014 |
| readonly-default | distinct | public-rule | 3,005,000 | 4,001 |
| readonly-default | distinct | complete-module-diagnostics | 3,015,006 | 14,007 |
| parentheses | repeated | public-rule | 1,003,000 | 1,003 |
| parentheses | repeated | complete-module-diagnostics | 1,011,015 | 9,018 |
| parentheses | distinct | public-rule | 1,502,500 | 3,001 |
| parentheses | distinct | complete-module-diagnostics | 1,517,508 | 18,009 |
| scalar-variable | repeated | public-rule | 1,004,001 | 2,004 |
| scalar-variable | repeated | complete-module-diagnostics | 1,011,014 | 9,017 |
| scalar-variable | distinct | public-rule | 1,503,501 | 4,002 |
| scalar-variable | distinct | complete-module-diagnostics | 1,516,508 | 17,009 |
| scalar-new | repeated | public-rule | 1,004,001 | 2,004 |
| scalar-new | repeated | complete-module-diagnostics | 1,009,007 | 7,010 |
| scalar-new | distinct | public-rule | 1,503,501 | 4,002 |
| scalar-new | distinct | complete-module-diagnostics | 1,509,506 | 10,007 |

These fixtures bound the repaired public-rule metadata work. They do not establish that every analyzer path has linear work.

## Timings

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor. Sequential baseline/current/current/baseline trials, three warmups and nine measured rounds. Ranges are the two trial medians in milliseconds. Public-rule timing excludes parsing/symbols and explicit priming; cold assignment timing includes building its object-state walks with a fresh context. Complete-module timing includes analyzer setup and internal-error capture. Getter counters are removed before timing; metadata uses plain arrays/properties. Independent expectations and complete equality checks run outside the clock. No editor, heap or cold-start claim.

| Mode | Types | Scope | N | Before median ms | After median ms |
| --- | --- | --- | ---: | ---: | ---: |
| cold-no-default | repeated | public-rule | 1 | 0.0511–0.0542 | 0.0481–0.0560 |
| cold-no-default | repeated | complete-module-diagnostics | 1 | 1.1777–1.2284 | 1.0603–1.1251 |
| cold-no-default | repeated | public-rule | 100 | 1.1040–1.1453 | 0.5341–0.6376 |
| cold-no-default | repeated | complete-module-diagnostics | 100 | 8.6595–8.8859 | 8.8903–9.6409 |
| cold-no-default | repeated | public-rule | 1000 | 66.1174–80.3125 | 7.0223–7.0662 |
| cold-no-default | repeated | complete-module-diagnostics | 1000 | 106.4524–111.3992 | 102.6440–105.1003 |
| cold-no-default | distinct | public-rule | 1 | 0.0040–0.0040 | 0.0036–0.0074 |
| cold-no-default | distinct | complete-module-diagnostics | 1 | 0.3209–0.3218 | 0.3360–0.3595 |
| cold-no-default | distinct | public-rule | 100 | 1.7456–1.8134 | 0.3807–0.4055 |
| cold-no-default | distinct | complete-module-diagnostics | 100 | 5.9185–8.0203 | 5.9785–7.0801 |
| cold-no-default | distinct | public-rule | 1000 | 136.6225–150.1647 | 7.5791–8.1304 |
| cold-no-default | distinct | complete-module-diagnostics | 1000 | 146.0517–151.1322 | 95.9884–103.9815 |
| primed-no-default | repeated | public-rule | 1 | 0.0016–0.0030 | 0.0017–0.0026 |
| primed-no-default | repeated | complete-module-diagnostics | 1 | 0.3668–0.3672 | 0.3245–0.3450 |
| primed-no-default | repeated | public-rule | 100 | 0.0704–0.0759 | 0.0715–0.0716 |
| primed-no-default | repeated | complete-module-diagnostics | 100 | 5.4608–6.2534 | 5.6532–5.6919 |
| primed-no-default | repeated | public-rule | 1000 | 1.0066–1.0580 | 1.2131–1.3741 |
| primed-no-default | repeated | complete-module-diagnostics | 1000 | 85.2575–87.2605 | 82.5747–85.5497 |
| primed-no-default | distinct | public-rule | 1 | 0.0016–0.0025 | 0.0016–0.0016 |
| primed-no-default | distinct | complete-module-diagnostics | 1 | 0.3771–0.3905 | 0.2854–0.3001 |
| primed-no-default | distinct | public-rule | 100 | 0.5491–0.5645 | 0.1447–0.1611 |
| primed-no-default | distinct | complete-module-diagnostics | 100 | 5.1712–6.1200 | 5.1359–5.7980 |
| primed-no-default | distinct | public-rule | 1000 | 47.7081–48.0668 | 1.8022–2.2533 |
| primed-no-default | distinct | complete-module-diagnostics | 1000 | 135.9626–137.9555 | 90.7995–90.9530 |
| writable-default | repeated | public-rule | 1 | 0.0033–0.0038 | 0.0031–0.0034 |
| writable-default | repeated | complete-module-diagnostics | 1 | 0.3033–0.3249 | 0.2907–0.3003 |
| writable-default | repeated | public-rule | 100 | 0.0771–0.0821 | 0.0796–0.0898 |
| writable-default | repeated | complete-module-diagnostics | 100 | 5.6267–5.9952 | 4.6699–5.1310 |
| writable-default | repeated | public-rule | 1000 | 0.7239–1.0089 | 0.7592–0.8852 |
| writable-default | repeated | complete-module-diagnostics | 1000 | 77.8473–84.1433 | 82.2926–83.4910 |
| writable-default | distinct | public-rule | 1 | 0.0021–0.0021 | 0.0025–0.0034 |
| writable-default | distinct | complete-module-diagnostics | 1 | 0.2702–0.2819 | 0.2755–0.4374 |
| writable-default | distinct | public-rule | 100 | 0.6774–0.7086 | 0.2110–0.2182 |
| writable-default | distinct | complete-module-diagnostics | 100 | 5.9582–7.3066 | 5.9305–6.0974 |
| writable-default | distinct | public-rule | 1000 | 58.7719–58.8050 | 2.5754–2.7075 |
| writable-default | distinct | complete-module-diagnostics | 1000 | 145.1533–147.4772 | 84.7323–90.6233 |
| readonly-default | repeated | public-rule | 1 | 0.0020–0.0021 | 0.0020–0.0021 |
| readonly-default | repeated | complete-module-diagnostics | 1 | 0.3665–0.4193 | 0.2682–0.2724 |
| readonly-default | repeated | public-rule | 100 | 0.0746–0.0799 | 0.0848–0.0863 |
| readonly-default | repeated | complete-module-diagnostics | 100 | 5.3253–5.3265 | 4.5835–4.7078 |
| readonly-default | repeated | public-rule | 1000 | 0.7475–0.7554 | 0.7534–0.7994 |
| readonly-default | repeated | complete-module-diagnostics | 1000 | 78.9923–87.2975 | 78.8203–81.8780 |
| readonly-default | distinct | public-rule | 1 | 0.0020–0.0022 | 0.0018–0.0021 |
| readonly-default | distinct | complete-module-diagnostics | 1 | 0.2768–0.3037 | 0.2579–0.2908 |
| readonly-default | distinct | public-rule | 100 | 0.6963–0.7344 | 0.1704–0.1931 |
| readonly-default | distinct | complete-module-diagnostics | 100 | 5.7524–6.0841 | 5.8385–6.0725 |
| readonly-default | distinct | public-rule | 1000 | 55.1711–55.7037 | 2.5128–2.5172 |
| readonly-default | distinct | complete-module-diagnostics | 1000 | 145.4325–145.7455 | 83.7347–87.7548 |
| parentheses | repeated | public-rule | 1 | 0.0023–0.0025 | 0.0036–0.0041 |
| parentheses | repeated | complete-module-diagnostics | 1 | 0.7259–0.8175 | 0.7473–0.9567 |
| parentheses | repeated | public-rule | 100 | 0.2580–0.3623 | 0.0532–0.0797 |
| parentheses | repeated | complete-module-diagnostics | 100 | 8.9356–9.9321 | 9.5906–10.1352 |
| parentheses | repeated | public-rule | 1000 | 19.7876–21.0724 | 0.3069–0.3099 |
| parentheses | repeated | complete-module-diagnostics | 1000 | 130.8425–132.5937 | 109.2990–115.5029 |
| parentheses | distinct | public-rule | 1 | 0.0008–0.0008 | 0.0011–0.0012 |
| parentheses | distinct | complete-module-diagnostics | 1 | 0.3570–0.4077 | 0.3574–0.3738 |
| parentheses | distinct | public-rule | 100 | 0.3246–0.5285 | 0.1055–0.1900 |
| parentheses | distinct | complete-module-diagnostics | 100 | 8.1204–8.1670 | 8.2531–8.7126 |
| parentheses | distinct | public-rule | 1000 | 28.2439–28.5087 | 1.1193–1.5473 |
| parentheses | distinct | complete-module-diagnostics | 1000 | 136.5069–136.8644 | 106.6917–116.5914 |
| scalar-variable | repeated | public-rule | 1 | 0.0199–0.0210 | 0.0161–0.0163 |
| scalar-variable | repeated | complete-module-diagnostics | 1 | 0.9634–0.9936 | 0.9292–1.0491 |
| scalar-variable | repeated | public-rule | 100 | 0.5827–0.7436 | 0.3400–0.3523 |
| scalar-variable | repeated | complete-module-diagnostics | 100 | 8.2894–8.3452 | 8.4135–10.2254 |
| scalar-variable | repeated | public-rule | 1000 | 22.7399–23.0210 | 2.7032–3.0746 |
| scalar-variable | repeated | complete-module-diagnostics | 1000 | 118.3956–146.0594 | 96.2942–102.7114 |
| scalar-variable | distinct | public-rule | 1 | 0.0073–0.0077 | 0.0077–0.0119 |
| scalar-variable | distinct | complete-module-diagnostics | 1 | 0.4709–0.5038 | 0.4515–0.5751 |
| scalar-variable | distinct | public-rule | 100 | 0.4986–0.7515 | 0.3008–0.4110 |
| scalar-variable | distinct | complete-module-diagnostics | 100 | 6.5186–7.2134 | 6.8241–7.7298 |
| scalar-variable | distinct | public-rule | 1000 | 30.1424–32.0095 | 3.3284–3.4026 |
| scalar-variable | distinct | complete-module-diagnostics | 1000 | 133.9833–135.2925 | 99.3827–99.5618 |
| scalar-new | repeated | public-rule | 1 | 0.0378–0.0386 | 0.0148–0.0220 |
| scalar-new | repeated | complete-module-diagnostics | 1 | 0.6505–0.8132 | 0.6523–0.6535 |
| scalar-new | repeated | public-rule | 100 | 0.4531–0.5631 | 0.2366–0.3315 |
| scalar-new | repeated | complete-module-diagnostics | 100 | 7.2373–9.3681 | 6.9048–7.2569 |
| scalar-new | repeated | public-rule | 1000 | 21.9745–23.4787 | 2.1270–2.3747 |
| scalar-new | repeated | complete-module-diagnostics | 1000 | 123.6680–129.7428 | 86.8880–86.9139 |
| scalar-new | distinct | public-rule | 1 | 0.0072–0.0077 | 0.0073–0.0078 |
| scalar-new | distinct | complete-module-diagnostics | 1 | 0.4077–0.5797 | 0.4126–0.4304 |
| scalar-new | distinct | public-rule | 100 | 0.5971–0.7722 | 0.3188–0.4126 |
| scalar-new | distinct | complete-module-diagnostics | 100 | 6.3380–8.5064 | 5.4600–5.5917 |
| scalar-new | distinct | public-rule | 1000 | 29.8941–30.8477 | 3.0734–3.1651 |
| scalar-new | distinct | complete-module-diagnostics | 1000 | 125.5986–129.9616 | 92.5735–97.8621 |

Every 1,000-procedure distinct-type public-rule and complete-module fixture improves in both trials. Repeated-type large results are mixed: primed no-default public assignments are slower (1.01–1.06 ms to 1.21–1.37 ms), and writable/read-only assignment ranges overlap. Repeated writable full-module trials also overlap and include a slower after trial. Several one-procedure cases overlap or are slower. At 100 procedures, repeated cold assignments, distinct parentheses and repeated scalar-variable full-module calls are slower in both trials; other cases are mixed or overlap. These measurements establish large worst-case gains and bounded metadata work, not a universal latency improvement.

Reproduce from the repository root, sequentially:

```powershell
node scripts/benchmark-remaining-default-member-queries.mjs --baseline=edb0d30923fe30ffaf5c3a67dd0b03b02b543b08
node scripts/benchmark-remaining-default-member-queries.mjs
node scripts/benchmark-remaining-default-member-queries.mjs
node scripts/benchmark-remaining-default-member-queries.mjs --baseline=edb0d30923fe30ffaf5c3a67dd0b03b02b543b08
```
