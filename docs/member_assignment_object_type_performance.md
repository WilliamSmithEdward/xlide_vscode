# Member assignment object-type query performance

Baseline: `2bad3627dde0fc20c77a73e0e372c75f4b564a4b` (PR #1195). Measurements used source head `5c540bff9c89dcefa862616df8dcba1ef95e062a`; its entire tracked tree was verified identical to the published merge. Baseline comparisons restore assignments.ts; all surrounding files are identical.

The member-assignment predicate resolves the same expected object type for every assignment, even though the public rule already owns a lazy type resolver. Pass that resolver through the private member-assignment checker and use the existing optional resolver argument to isKnownObjectAssignmentType. The three-line change preserves setter, array, writable and eligibility guards, query freshness and diagnostic text/spans. No dependency or cache lifetime is added.

## Validation

Types and 153 focused tests pass. Twenty-four new tests include six class-name work bounds at 10/100/1,000 assignments with repeated/distinct expected types; six surface-kind eligibility controls; seven generic/host/scalar/unknown/case controls; three Let/array/writable guards; ambiguous eligible names; retained-AST metadata freshness. The baseline fails four work bounds and passes two small work cases and all 18 semantic controls. All work tests assert complete exact Set-required messages and member spans.

Full suite: 15,722 passed, 33 skipped across 797 passing files and seven skipped files, in 79.87 seconds. Complete outputs match 20,000 generated four-rule bundles, 2,000 generated complete modules and 24,768 corpus diagnostic/internal-error comparisons over 8,256 sources with LF/CRLF/CR; zero differences. Generated cases vary member return/write types, writable/Let/Set flags, class kinds/completeness/duplicates, generic/scalar/project/host/unknown types and member assignments with scalar/New/Set values, alongside existing default-value/argument/block forms. No diagnostic changes are intended.

## Work

Fixtures use N member assignments on an As New Holder. Repeated types write Item As Class0; distinct types write Item0..ItemN-1 with corresponding Class types. Metadata has N+2 class surfaces. Class-name and member-name getters run only outside timing. Independent expected output is exactly N Set-required messages and member spans, with no full-module internal errors.

| N = 1,000 types | Scope | Before class-name reads | After class-name reads | Before member-name reads | After member-name reads |
| --- | --- | ---: | ---: | ---: | ---: |
| repeated | public-rule | 1,005,002 | 3,005 | 5,000 | 5,000 |
| repeated | complete-module-diagnostics | 1,008,014 | 6,017 | 8,003 | 8,003 |
| distinct | public-rule | 1,005,002 | 4,004 | 2,502,500 | 2,502,500 |
| distinct | complete-module-diagnostics | 1,008,014 | 7,016 | 1,509,500 | 1,509,500 |

This bounds the repaired class-name query in these public-rule fixtures. Member lookup and other analyzer paths are separate; no universal linear-work claim is made.

## Timings

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor. Sequential baseline/current/current/baseline trials, three warmups and nine measured rounds. Ranges are the two trial medians in milliseconds. Public-rule timing excludes parsing/symbol construction; complete-module timing includes analyzer setup and internal-error capture. All getters are replaced with plain properties before timing. Independent expectations and complete equality checks run outside the clock. No editor, heap or cold-start claim.

| Types | Scope | N | Before median ms | After median ms |
| --- | --- | ---: | ---: | ---: |
| repeated | public-rule | 1 | 0.0328–0.0331 | 0.0327–0.0343 |
| repeated | complete-module-diagnostics | 1 | 1.0812–1.1532 | 1.0670–1.0843 |
| repeated | public-rule | 100 | 1.0437–1.1015 | 0.7640–0.7894 |
| repeated | complete-module-diagnostics | 100 | 4.1842–4.2994 | 4.0063–4.0390 |
| repeated | public-rule | 1000 | 25.7593–26.7485 | 7.2966–7.5992 |
| repeated | complete-module-diagnostics | 1000 | 40.6029–40.8728 | 23.7811–24.1555 |
| distinct | public-rule | 1 | 0.0033–0.0034 | 0.0036–0.0036 |
| distinct | complete-module-diagnostics | 1 | 0.4940–0.5059 | 0.5196–0.5261 |
| distinct | public-rule | 100 | 1.6233–1.6559 | 1.5487–1.5613 |
| distinct | complete-module-diagnostics | 100 | 3.1219–3.1759 | 3.0049–3.0189 |
| distinct | public-rule | 1000 | 137.8167–139.2653 | 117.4968–122.9502 |
| distinct | complete-module-diagnostics | 1000 | 60.8200–61.9845 | 43.6309–45.8509 |

Both 1,000-assignment fixtures improve in both trials. At one distinct assignment, both public-rule and complete-module medians are slightly slower; repeated single-assignment ranges overlap. These results do not establish an improvement for every small module. Distinct-member lookup still performs 2,502,500 public-rule and 1,509,500 complete-module member-name reads; this repair removes the separate class-name scan and leaves that confirmed residual for further attribution.

Reproduce from the repository root, sequentially:

```powershell
node scripts/benchmark-member-assignment-object-types.mjs --baseline=2bad3627dde0fc20c77a73e0e372c75f4b564a4b
node scripts/benchmark-member-assignment-object-types.mjs
node scripts/benchmark-member-assignment-object-types.mjs
node scripts/benchmark-member-assignment-object-types.mjs --baseline=2bad3627dde0fc20c77a73e0e372c75f4b564a4b
```
