# Class-instance member lookup

The class-instance rule used a case-insensitive Array.find for every member read. With N members and N reads, this produced quadratic member-name work even after statements were indexed.

Use a query-owned member index built only when a tracked receiver actually reads a member. Preserve the first duplicate name, including case-insensitive duplicates. Share the consulted index across procedures; rebuild it on every public rule query so project metadata replacement remains visible. Classes with at most eight members retain a bounded linear scan.

## Reproduction

Run `node scripts/benchmark-class-member-index.mjs`, then with `--baseline=1d64d8500f6e338d9c94944e03a84a7c9c2c8edb`. The baseline overrides only this rule. Each fixture has one local `As New Class1`, N metadata members and N `Debug.Print actor.M.Count` reads. Repeated fixtures read the last member; distinct fixtures read every member once. Members independently hold Nothing.

Sequential before/after/after/before runs on Node 24.18.0 and Ryzen 7 9800X3D. Three warmups, nine measured calls, ranges of the two run medians in milliseconds.

| Reads | Members | Scope | Before ms | After ms |
| --- | ---: | --- | ---: | ---: |
| repeated | 1 | rule | 0.0058–0.0072 | 0.0057–0.006 |
| repeated | 1 | complete-module-diagnostics | 1.1411–1.2011 | 1.152–1.1523 |
| repeated | 100 | rule | 0.237–0.2479 | 0.103–0.1282 |
| repeated | 100 | complete-module-diagnostics | 6.209–6.3933 | 5.9288–6.3615 |
| repeated | 1000 | rule | 16.6031–16.9308 | 0.4535–0.4567 |
| repeated | 1000 | complete-module-diagnostics | 94.7537–96.0837 | 77.0083–77.8182 |
| distinct | 1 | rule | 0.0011–0.0012 | 0.0015–0.0016 |
| distinct | 1 | complete-module-diagnostics | 0.5288–0.5334 | 0.5315–0.5326 |
| distinct | 100 | rule | 0.1189–0.1198 | 0.0509–0.0566 |
| distinct | 100 | complete-module-diagnostics | 3.9847–4.2361 | 3.8708–3.9378 |
| distinct | 1000 | rule | 8.0944–8.1319 | 0.4609–0.4718 |
| distinct | 1000 | complete-module-diagnostics | 60.3485–61.3845 | 51.8853–52.4623 |

At 1,000 members/reads, repeated-read member-name accesses fall from 1,001,000 to 2,000; distinct reads from 501,500 to 2,000. Counts include the N diagnostic display-name accesses. The direct rule independently checks every exact message and original span; complete-module calls independently require N object-variable-not-set diagnostics and no internal failures, then freeze and compare complete output each round.

Small results are mixed, and the repeated 100-member complete-call ranges overlap. AST/symbol preparation is outside direct-rule timing. Complete-module timings include analyzer preparation and no-error assertions; equality checks are outside timing. No editor, cold-start or heap improvement is claimed. A one-off query touching one member of a large class now pays full index construction, a deliberate tradeoff for bounding repeated and distinct lookups.

## Validation

- Type checking and 35 focused tests pass.
- Eleven new tests: seven fail on the old rule, four semantic controls already pass. Independent work bounds cover repeated/distinct reads and sharing between procedures; controls cover tiny/indexed duplicate selection, mutable metadata between queries and lazy access.
- Full suite: 15,299 tests across 781 files pass, with 33 tests and seven files skipped.
- 20,000 generated complete rule results and 24,768 full module diagnostic/error results match the baseline. Generated cases use large surfaces with case-insensitive duplicates and vary the first member fact between Nothing, Empty, scalar and unknown across successive queries. Oracle comparisons cover 8,256 sources and LF/CRLF/CR.
