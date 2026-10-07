# TypeOf project metadata queries

For each TypeOf condition, the rule repeatedly resolved operand/target object types, scanned every supplied Implements list to establish whether the operand could be an interface, and scanned the project twice more for shared implementers. With N conditions and N project surfaces, these scans were quadratic.

Reuse the existing lazy createObjectAssignmentTypeResolver and createProjectInterfaceSharingLookup in both compatibility directions. Lazily index the concrete-operand interface exclusion once. All derived data belongs to one public rule query; later queries observe changed metadata. Preserve every supplied surface kind for interface exclusion, eligible duplicate-name ambiguity, host/library priority and original display text/spans. Remove the superseded private scanner.

## Reproduction and measurements

Run `node scripts/benchmark-typeof-query-metadata.mjs`, then `--baseline=ad88862311cc7b68d9270b8a4b965899c83e7ae7`. Only the TypeOf rule is overridden for baseline. Fixtures have N parsed TypeOf conditions, N+1 plain project class surfaces and no Implements memberships. Repeated fixtures compare Class1 to Class2; distinct fixtures compare it to each other class. Every condition independently requires one exact always-false diagnostic.

Work counters use getters for one untimed rule call. The harness replaces them with plain name/implements properties before either timed scope, so counter increments and dynamically constructed names do not inflate timing. Sequential before/after/after/before on Node 24.18.0, Ryzen 7 9800X3D, three warmups and nine measured calls. Ranges of the two run medians, in milliseconds:

| Targets | Conditions | Scope | Before ms | After ms |
| --- | ---: | --- | ---: | ---: |
| repeated | 1 | rule | 0.0143–0.0143 | 0.0122–0.0122 |
| repeated | 1 | complete-module-diagnostics | 1.1792–1.3066 | 1.2043–1.3194 |
| repeated | 100 | rule | 2.1885–2.3241 | 0.1407–0.1416 |
| repeated | 100 | complete-module-diagnostics | 8.5893–8.9095 | 6.9078–7.1986 |
| repeated | 1000 | rule | 163.2309–169.9867 | 1.0916–1.1742 |
| repeated | 1000 | complete-module-diagnostics | 222.6658–225.6183 | 43.7704–44.8205 |
| distinct | 1 | rule | 0.0031–0.0033 | 0.0022–0.0022 |
| distinct | 1 | complete-module-diagnostics | 0.5344–0.5417 | 0.504–0.5345 |
| distinct | 100 | rule | 1.8903–1.983 | 0.1621–0.17 |
| distinct | 100 | complete-module-diagnostics | 6.5298–6.8465 | 4.4326–5.1845 |
| distinct | 1000 | rule | 171.6018–172.1246 | 1.6488–1.7743 |
| distinct | 1000 | complete-module-diagnostics | 222.3836–234.6631 | 51.0134–51.3596 |

At 1,000 conditions, repeated-target name accesses fall from 6,012,000 to 1,003 and Implements accesses from 3,009,000 to 2,004. For distinct targets they fall to 2,002 and 3,003 respectively. Independent tests bound both counters by 4N+12 while asserting every exact diagnostic and span.

Tiny full-call ranges overlap; no general small-query speedup is claimed. A single interface exclusion now pays index construction across the project even if the old scan could stop early. AST/symbol preparation is outside direct-rule timing. Complete module calls include analyzer setup and no-error assertions; complete-output equality checks are outside timing. No editor, cold-start or heap claim is made.

## Validation

- Types and 51 focused tests pass.
- Twenty new tests: six baseline work failures and 14 controls already passing on baseline. Coverage includes repeated/distinct targets, duplicate ambiguity, excluded value-type duplicates, all surface kinds, direct implementation compatibility, changed metadata on later queries and generic lazy handling.
- Full suite: 15,324 tests across 783 files pass, with 33 tests and seven files skipped.
- Complete baseline/current outputs match for 20,000 generated rule cases with varying types/surface kinds/implements/duplicate names, 2,000 generated full module cases and 24,768 full module diagnostic/error results over 8,256 oracle sources across LF/CRLF/CR.
