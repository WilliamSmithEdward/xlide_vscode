# Late-bound class metadata

The runtime-member rule rebuilt member-name, accessor, scalar-field and method-parameter metadata whenever a late-bound receiver was assigned a known project class or a known Collection item was read. N references to a class with N members therefore caused quadratic member work. Distinct class names also repeatedly searched the entire project surface.

Use one lazy class lookup for the complete public module query, shared between procedures, typed assignments and collection reads. Index only exhaustive plain classes and retain the first eligible case-insensitive duplicate, matching the former find. Project member metadata is built only for consulted classes. Keep per-receiver mayBeNothing state on copied overlays, and rebuild the lookup on the next public query so changed metadata is visible. Collection still bypasses project indexing; unknown classes stay unknown.

## Reproduction

Run `node scripts/benchmark-late-bound-class-metadata.mjs`, then with `--baseline=0f527d3e4753a974b5869128bd1e7f08f09b489f`. Only lateBoundMembers.ts is overridden by the pinned baseline. All fixtures use N separate procedures and N+1 project classes. Repeated direct calls assign an Object to New Class1 then read actor.Nope; repeated Collection calls add New Class1 and read c(1).Nope. Class1 has N scalar fields. Distinct direct calls use Class1 through ClassN, each with one scalar field.

Direct queries independently check N exact runtime-member diagnostics, including messages and original spans. Complete-module calls independently require exactly N runtime-member-not-found diagnostics and no internal failures, then compare complete output every round. Counters are exclusively untimed; all class/member names become plain string properties before timing.

Sequential before/after/after/before runs on Node 24.18.0 and Ryzen 7 9800X3D. Three warmups, nine measured calls, ranges of the two run medians in milliseconds. Parsed AST and symbols are outside direct-rule timing; complete-module calls include analyzer preparation and error checks. Output equality is outside timing.

| Fixture | Procedures | Scope | Before ms | After ms |
| --- | ---: | --- | ---: | ---: |
| repeated-direct | 1 | rule | 0.0977–0.098 | 0.0895–0.1056 |
| repeated-direct | 1 | complete-module-diagnostics | 1.3752–1.5076 | 1.3169–1.3697 |
| repeated-direct | 100 | rule | 2.9121–3.1861 | 1.0242–1.0546 |
| repeated-direct | 100 | complete-module-diagnostics | 16.0247–19.1339 | 13.4727–13.6764 |
| repeated-direct | 1000 | rule | 173.0597–175.8801 | 8.1726–9.6511 |
| repeated-direct | 1000 | complete-module-diagnostics | 327.0909–334.4983 | 139.0843–142.8983 |
| repeated-collection | 1 | rule | 0.0512–0.0891 | 0.0475–0.0485 |
| repeated-collection | 1 | complete-module-diagnostics | 0.8322–0.9643 | 0.8083–0.8089 |
| repeated-collection | 100 | rule | 2.4919–3.0465 | 0.9525–0.9803 |
| repeated-collection | 100 | complete-module-diagnostics | 13.5856–19.8739 | 12.5519–18.2283 |
| repeated-collection | 1000 | rule | 179.8221–183.2097 | 10.6192–12.6322 |
| repeated-collection | 1000 | complete-module-diagnostics | 308.6217–308.6728 | 151.8447–169.9948 |
| distinct-direct | 1 | rule | 0.0367–0.0368 | 0.0369–0.0652 |
| distinct-direct | 1 | complete-module-diagnostics | 0.4652–0.4661 | 0.5371–0.5498 |
| distinct-direct | 100 | rule | 0.8695–0.9855 | 1.3805–1.4072 |
| distinct-direct | 100 | complete-module-diagnostics | 9.6459–9.8572 | 10.1394–13.6674 |
| distinct-direct | 1000 | rule | 18.6438–18.7086 | 9.4409–12.1458 |
| distinct-direct | 1000 | complete-module-diagnostics | 133.1579–139.8475 | 126.9085–181.5718 |

At 1,000 repeated direct or collection reads, public-query member-name work falls from 2,000,000 to 2,000, and project-name work from 3,001 to 2,003. Distinct-type project-name work falls from 502,501 to 3,002; its 2,000 member-name reads are unchanged. Counts include other unchanged work within the public query.

Small and distinct-type timings are mixed. The repeated one-procedure direct-rule ranges overlap. The distinct one-procedure rule and complete calls and distinct 100-procedure rule and complete calls are slower. Repeated 100-procedure Collection complete ranges overlap; distinct 1,000-procedure complete ranges overlap widely, with one after median slower. These measurements establish the repeated-class benefit and bounded work, not a uniform improvement in full diagnostics. A single class query now pays an index over all eligible class names instead of an early find, while member projection remains lazy. No editor, cold-start or heap improvement is claimed.

## Validation

- Types and 37 focused tests pass.
- Twenty-four new tests: thirteen work-bound failures on baseline and eleven already-passing semantic controls. Tests independently check exact messages/spans across direct New, typed assignments, Collection reads and distinct types; shared projection between consumers; first eligible duplicates; excluded/incomplete/unused surfaces; lazy Collection and inactive paths; fresh metadata; isolated Nothing overlays.
- Full suite: 15,356 tests across 785 files pass; 33 tests and seven files skipped.
- 20,000 generated complete public-rule results, 2,000 generated complete-module results and 24,768 oracle module diagnostic/error results match baseline with no differences. Generated cases vary member kind, accessor flags, signatures, scalar/object returns, duplicate names, completeness, typed/auto-instanced/New/ProgID receivers, Collection reads, calls, assignments and conditional branches. Oracle comparison covers 8,256 sources across LF/CRLF/CR.
