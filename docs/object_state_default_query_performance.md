# Object-state default-member query performance

Baseline: `630a4def85bd7a7833476a73474d8b22a06fa1df` (PR #1191), including the helper from PR #1189. Measurements used audit merge `54c310cb4d5d232fe8bcedbb3eeefbea9bb1f6d8`; its entire tracked tree was verified identical to that published merge. Baseline comparisons restore objectState.ts; all surrounding files are identical.

Object-state statements repeatedly resolve the same object type and default-member facts for Let targets, conditions, value/operand reads and indexed reads. Create one lazy default-query bundle per public object-state invocation and pass it through procedure walks and every statement/header callback. Its resolver also supplies the existing classification paths. Direct Let-state calls create a default-query bundle only after their existing walk-cache check. All six verdict consumers and the set-object index-requirement consumer share the bundle. Source/module/symbol/activity/member-context cache ownership and freshness checks are unchanged. No dependency is added.

## Validation

Types and 130 focused tests pass. Fifty-three new tests include 27 public-rule name-read bounds at 10/100/1,000 reads across Let/condition/indexed paths and one-procedure/repeated-type/distinct-type procedure layouts. Twenty-two exact-finding/absent-default controls cover Let, plain value, operand, indexed/Set-indexed, single-line If, If block, While, Do entry/exit and Select headers. Four more controls cover refreshed metadata on a retained AST, incomplete-surface Let behavior, lazy unused member arrays and a direct Let-state call before the public rule creates its walk. The baseline fails all 27 work bounds and passes all 26 semantic/cache controls.

Full suite: 15,659 passed, 33 skipped across 795 passing files and seven skipped files, in 75.67 seconds. Complete outputs match 20,000 generated public-rule results, 2,000 generated full-module results and 24,768 corpus diagnostic/internal-error comparisons over 8,256 sources with LF/CRLF/CR; zero differences. Generated cases vary default-member signatures, class kinds/completeness/duplicates, local/module/array/function/parameter forms, scalar/generic/project/host/unknown types, Let/value/operator/index/coercion reads, guards and block headers. No diagnostic changes are intended.

## Work

Each fixture has N reads and N+1 class surfaces. One-procedure layouts repeat a read of Class0; procedure layouts declare Class0 or a distinct Class per procedure. The class has no default member, so the object-state rule independently emits no findings. Complete-module expectations assert exact code/message/span tuples for Set-required assignments and assigned-but-unread variables, empty condition findings and indexed no-default findings, plus no internal errors. Getter counters are untimed.

| N = 1,000 path | Layout | Scope | Before class-name reads | After class-name reads |
| --- | --- | --- | ---: | ---: |
| let | one-procedure | public-rule | 1,004,002 | 1,003 |
| let | one-procedure | complete-module-diagnostics | 1,011,015 | 8,016 |
| let | repeated-types | public-rule | 1,004,002 | 1,003 |
| let | repeated-types | complete-module-diagnostics | 1,011,015 | 8,016 |
| let | distinct-types | public-rule | 1,504,501 | 3,001 |
| let | distinct-types | complete-module-diagnostics | 4,017,006 | 2,515,506 |
| condition | one-procedure | public-rule | 2,007,002 | 1,003 |
| condition | one-procedure | complete-module-diagnostics | 2,014,014 | 8,015 |
| condition | repeated-types | public-rule | 2,007,002 | 1,003 |
| condition | repeated-types | complete-module-diagnostics | 2,014,014 | 8,015 |
| condition | distinct-types | public-rule | 3,007,001 | 3,001 |
| condition | distinct-types | complete-module-diagnostics | 3,019,008 | 15,008 |
| indexed | one-procedure | public-rule | 1,004,002 | 1,003 |
| indexed | one-procedure | complete-module-diagnostics | 1,012,011 | 9,012 |
| indexed | repeated-types | public-rule | 1,004,002 | 1,003 |
| indexed | repeated-types | complete-module-diagnostics | 1,012,011 | 9,012 |
| indexed | distinct-types | public-rule | 1,504,501 | 3,001 |
| indexed | distinct-types | complete-module-diagnostics | 1,515,507 | 14,007 |

Public-rule default-member name work is linear in these fixtures. Other rules still have default-member consumers; this does not establish a universal full-analyzer linear-work bound.

## Timings

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor. Sequential baseline/current/current/baseline trials, three warmups and nine measured rounds. Ranges are the two trial medians in milliseconds. Public-rule timing excludes parsing/symbol construction; complete-module timing includes analyzer setup and internal-error capture. Getters are replaced with plain properties before timing. Independent expected tuples and complete-output equality checks run outside the clock. No editor, cold-start or heap claim.

| Path | Layout | Scope | N | Before median ms | After median ms |
| --- | --- | --- | ---: | ---: | ---: |
| let | one-procedure | public-rule | 1 | 0.0441–0.0459 | 0.0424–0.0440 |
| let | one-procedure | complete-module-diagnostics | 1 | 1.1197–1.1270 | 0.9864–1.0031 |
| let | one-procedure | public-rule | 100 | 0.4111–0.4164 | 0.2964–0.3007 |
| let | one-procedure | complete-module-diagnostics | 100 | 3.3729–3.9020 | 3.2825–3.3451 |
| let | one-procedure | public-rule | 1000 | 18.3899–18.9018 | 1.1990–1.2719 |
| let | one-procedure | complete-module-diagnostics | 1000 | 35.6883–37.7993 | 18.8664–18.9212 |
| let | repeated-types | public-rule | 1 | 0.0044–0.0045 | 0.0046–0.0047 |
| let | repeated-types | complete-module-diagnostics | 1 | 0.4776–0.4784 | 0.4655–0.5162 |
| let | repeated-types | public-rule | 100 | 0.4900–0.5276 | 0.2964–0.3080 |
| let | repeated-types | complete-module-diagnostics | 100 | 6.3036–6.7269 | 7.2889–7.3535 |
| let | repeated-types | public-rule | 1000 | 21.7079–24.7294 | 4.5391–4.8408 |
| let | repeated-types | complete-module-diagnostics | 1000 | 95.2016–97.1046 | 77.1667–77.3045 |
| let | distinct-types | public-rule | 1 | 0.0027–0.0028 | 0.0028–0.0029 |
| let | distinct-types | complete-module-diagnostics | 1 | 0.3012–0.3164 | 0.3029–0.3058 |
| let | distinct-types | public-rule | 100 | 0.5659–0.5700 | 0.2951–0.2971 |
| let | distinct-types | complete-module-diagnostics | 100 | 5.7315–5.8443 | 5.4590–5.6125 |
| let | distinct-types | public-rule | 1000 | 27.9820–29.0001 | 5.2755–5.2912 |
| let | distinct-types | complete-module-diagnostics | 1000 | 139.2064–140.6748 | 114.7116–116.8364 |
| condition | one-procedure | public-rule | 1 | 0.0150–0.0150 | 0.0154–0.0263 |
| condition | one-procedure | complete-module-diagnostics | 1 | 0.7285–0.7408 | 0.7343–0.7468 |
| condition | one-procedure | public-rule | 100 | 0.7816–0.7935 | 0.3683–0.3887 |
| condition | one-procedure | complete-module-diagnostics | 100 | 5.1121–5.2507 | 4.9893–5.7850 |
| condition | one-procedure | public-rule | 1000 | 36.1451–36.8578 | 2.6911–2.7098 |
| condition | one-procedure | complete-module-diagnostics | 1000 | 70.4817–71.2577 | 35.6338–35.8265 |
| condition | repeated-types | public-rule | 1 | 0.0051–0.0053 | 0.0045–0.0045 |
| condition | repeated-types | complete-module-diagnostics | 1 | 0.4099–0.4385 | 0.3921–0.4033 |
| condition | repeated-types | public-rule | 100 | 0.7949–0.7980 | 0.3990–0.4179 |
| condition | repeated-types | complete-module-diagnostics | 100 | 7.7344–7.9202 | 6.5019–6.6071 |
| condition | repeated-types | public-rule | 1000 | 42.5512–43.5095 | 6.2130–6.2346 |
| condition | repeated-types | complete-module-diagnostics | 1000 | 130.0519–133.1951 | 90.6896–92.5408 |
| condition | distinct-types | public-rule | 1 | 0.0046–0.0046 | 0.0040–0.0042 |
| condition | distinct-types | complete-module-diagnostics | 1 | 0.3389–0.3461 | 0.3291–0.3348 |
| condition | distinct-types | public-rule | 100 | 0.9685–1.0631 | 0.4528–0.4632 |
| condition | distinct-types | complete-module-diagnostics | 100 | 7.1325–7.3347 | 6.8477–7.1458 |
| condition | distinct-types | public-rule | 1000 | 55.8049–56.9976 | 7.5305–7.5533 |
| condition | distinct-types | complete-module-diagnostics | 1000 | 138.4020–140.8642 | 90.9811–91.5311 |
| indexed | one-procedure | public-rule | 1 | 0.0064–0.0065 | 0.0063–0.0063 |
| indexed | one-procedure | complete-module-diagnostics | 1 | 0.6352–0.6610 | 0.6293–0.6757 |
| indexed | one-procedure | public-rule | 100 | 0.4049–0.4174 | 0.2084–0.2176 |
| indexed | one-procedure | complete-module-diagnostics | 100 | 4.1993–4.4698 | 3.9891–4.0220 |
| indexed | one-procedure | public-rule | 1000 | 18.5956–18.7296 | 1.8385–1.8390 |
| indexed | one-procedure | complete-module-diagnostics | 1000 | 47.3411–47.5289 | 29.7296–31.2584 |
| indexed | repeated-types | public-rule | 1 | 0.0039–0.0039 | 0.0036–0.0036 |
| indexed | repeated-types | complete-module-diagnostics | 1 | 0.4329–0.4377 | 0.3995–0.4357 |
| indexed | repeated-types | public-rule | 100 | 0.5023–0.5595 | 0.2985–0.3140 |
| indexed | repeated-types | complete-module-diagnostics | 100 | 6.2756–6.9794 | 6.0866–6.4758 |
| indexed | repeated-types | public-rule | 1000 | 22.5105–22.7088 | 5.2969–5.3485 |
| indexed | repeated-types | complete-module-diagnostics | 1000 | 106.1301–109.0125 | 87.6333–92.1554 |
| indexed | distinct-types | public-rule | 1 | 0.0035–0.0036 | 0.0033–0.0033 |
| indexed | distinct-types | complete-module-diagnostics | 1 | 0.3333–0.5461 | 0.3268–0.3484 |
| indexed | distinct-types | public-rule | 100 | 0.6185–0.8370 | 0.3487–0.3735 |
| indexed | distinct-types | complete-module-diagnostics | 100 | 7.0330–10.7559 | 6.1046–6.5238 |
| indexed | distinct-types | public-rule | 1000 | 30.2554–31.2753 | 6.3166–6.3518 |
| indexed | distinct-types | complete-module-diagnostics | 1000 | 108.8469–111.8920 | 84.6590–85.9190 |

Small cases are mixed: repeated/distinct one-read Let public-rule trials and the one-procedure condition public-rule trials are slower; several one-read full-module ranges overlap. The 100-procedure repeated-type Let full-module case is slower in both after trials (7.29–7.35 ms versus 6.30–6.73 ms). The one-procedure 100-condition full-module case includes a slower after trial. All 1,000-read public-rule and full-module cases improve in both trials. Distinct Let full-module work retains 2,515,506 name reads from other consumers; this repair does not remove that separately measured residual.

Reproduce from the repository root, sequentially:

```powershell
node scripts/benchmark-object-state-default-queries.mjs --baseline=630a4def85bd7a7833476a73474d8b22a06fa1df
node scripts/benchmark-object-state-default-queries.mjs
node scripts/benchmark-object-state-default-queries.mjs
node scripts/benchmark-object-state-default-queries.mjs --baseline=630a4def85bd7a7833476a73474d8b22a06fa1df
```
