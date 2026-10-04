# Class-instance member lookup

The class-instance rule used a case-insensitive Array.find for every member read. With N members and N reads, this produced quadratic member-name work even after statements were indexed.

Use a query-owned member index built only when a tracked receiver actually reads a member. Preserve the first duplicate name, including case-insensitive duplicates. Share the consulted index across procedures; rebuild it on every public rule query so project metadata replacement remains visible. Classes with at most eight members retain a bounded linear scan.

## Reproduction

Run `node scripts/benchmark-class-member-index.mjs`, then with `--baseline=457df8053a79ef1e7b30117701ac1481a04a7a23`. The baseline overrides only this rule. Each fixture has one local `As New Class1`, N metadata members and N `Debug.Print actor.M.Count` reads. Repeated fixtures read the last member; distinct fixtures read every member once. Members independently hold Nothing.

The previous timing table used name getters that incremented a counter and constructed `M` + index during timed reads. That overhead inflated the comparison against production plain metadata. The corrected table below replaces those timings; the deterministic work counts and production-rule correctness checks remain valid. Getter instrumentation now runs only in the separate untimed work check, and metadata names become plain properties before both timed scopes.

Sequential before/after/after/before runs on Node 24.18.0 and Ryzen 7 9800X3D. Three warmups, nine measured calls, ranges of the two run medians in milliseconds.

| Reads | Members | Scope | Before ms | After ms |
| --- | ---: | --- | ---: | ---: |
| repeated | 1 | rule | 0.0059–0.0077 | 0.0061–0.0073 |
| repeated | 1 | complete-module-diagnostics | 1.1657–1.4783 | 1.1817–1.186 |
| repeated | 100 | rule | 0.1821–0.1853 | 0.1083–0.1208 |
| repeated | 100 | complete-module-diagnostics | 6.1665–6.4585 | 5.8557–6.0317 |
| repeated | 1000 | rule | 8.8664–10.9849 | 0.4596–0.4754 |
| repeated | 1000 | complete-module-diagnostics | 71.0838–73.2427 | 63.5478–65.464 |
| distinct | 1 | rule | 0.0011–0.0011 | 0.0016–0.0027 |
| distinct | 1 | complete-module-diagnostics | 0.5298–0.5341 | 0.5843–0.669 |
| distinct | 100 | rule | 0.0723–0.0772 | 0.0511–0.0661 |
| distinct | 100 | complete-module-diagnostics | 3.8353–3.8807 | 3.8475–4.2197 |
| distinct | 1000 | rule | 4.3615–4.3778 | 0.4444–0.8101 |
| distinct | 1000 | complete-module-diagnostics | 49.8348–50.0499 | 44.499–46.979 |

At 1,000 members/reads, repeated-read member-name accesses fall from 1,001,000 to 2,000; distinct reads from 501,500 to 2,000. Counts include the N diagnostic display-name accesses. The direct rule independently checks every exact message and original span; complete-module calls independently require N object-variable-not-set diagnostics and no internal failures, then freeze and compare complete output each round.

Small results are mixed: the distinct one-member direct and complete calls are slower, the repeated one-member ranges overlap, and the distinct 100-member complete-call ranges overlap. AST/symbol preparation is outside direct-rule timing. Complete-module timings include analyzer preparation and no-error assertions; equality checks are outside timing. No editor, cold-start or heap improvement is claimed. A one-off query touching one member of a large class now pays full index construction, a deliberate tradeoff for bounding repeated and distinct lookups.

## Validation

- Type checking and 35 focused tests pass.
- Eleven new tests: seven fail on the old rule, four semantic controls already pass. Independent work bounds cover repeated/distinct reads and sharing between procedures; controls cover tiny/indexed duplicate selection, mutable metadata between queries and lazy access.
- Full suite: 15,299 tests across 781 files pass, with 33 tests and seven files skipped.
- 20,000 generated complete rule results and 24,768 full module diagnostic/error results match the baseline. Generated cases use large surfaces with case-insensitive duplicates and vary the first member fact between Nothing, Empty, scalar and unknown across successive queries. Oracle comparisons cover 8,256 sources and LF/CRLF/CR.

The benchmark-only correction reruns all repeated/distinct fixtures with three warmups and nine measured calls in ABBA order. Exact direct-rule diagnostics/spans and independent complete-module count/code/no-error controls pass in every run. Runtime code and regression tests are unchanged; the validation counts above describe the original member-index repair.
