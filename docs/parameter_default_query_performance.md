# Parameter-default project-type queries

Both public parameter-default rules resolved object types separately for every Optional default. Each project resolution filtered all eligible project types, even when the same type was used across many procedures. N parameters against N types therefore caused quadratic name work in both the compatibility and nonconstant-default checks.

Each rule now owns one existing lazy object-type resolver for its complete module query. It preserves host/library priority, generic and scalar handling, eligible duplicate ambiguity, excluded value-type/module kinds, display spelling, original spans and conditional activity. A subsequent public query constructs a fresh resolver and observes changed project metadata.

## Reproduction

Run `node scripts/benchmark-parameter-default-queries.mjs`, then with `--baseline=3fd41533df8e4dfee5046acf5a9a240883d8f775`. Only declarations.ts is overridden by the pinned baseline. Each fixture has N separate procedures, N+1 project class surfaces and two Optional parameters per procedure: an actor defaulting to 0, and a Long defaulting to Factory(). Repeated fixtures use Class1 throughout; distinct fixtures use Class1 through ClassN.

Each direct public rule independently checks N exact diagnostics, including messages and original spans. Complete-module calls independently require N parameter-default-type-mismatch and N parameter-default-not-constant diagnostics and no internal failures, then compare complete output every round. Counters run only in separate untimed checks; names are plain string properties during both timed scopes.

Sequential before/after/after/before runs on Node 24.18.0 and Ryzen 7 9800X3D. Three warmups, nine measured calls, ranges of the two run medians in milliseconds. Parsed AST preparation is outside direct-rule timing; complete-module calls include analyzer preparation and error checks. Output equality is outside timing.

| Types | Procedures | Scope | Before ms | After ms |
| --- | ---: | --- | ---: | ---: |
| repeated | 1 | checkParameterDefaultValues | 0.0135–0.0141 | 0.0132–0.0133 |
| repeated | 1 | checkNonConstantParameterDefaults | 0.0034–0.0035 | 0.0037–0.0038 |
| repeated | 1 | complete-module-diagnostics | 0.6892–0.8935 | 0.6555–0.6754 |
| repeated | 100 | checkParameterDefaultValues | 0.3371–0.3502 | 0.1463–0.147 |
| repeated | 100 | checkNonConstantParameterDefaults | 0.2462–0.2531 | 0.0311–0.0386 |
| repeated | 100 | complete-module-diagnostics | 6.2942–6.4382 | 5.9537–6.0496 |
| repeated | 1000 | checkParameterDefaultValues | 18.6504–18.7129 | 1.0303–1.0421 |
| repeated | 1000 | checkNonConstantParameterDefaults | 17.8752–18.0693 | 0.2489–0.2756 |
| repeated | 1000 | complete-module-diagnostics | 109.186–110.7708 | 72.3839–74.6201 |
| distinct | 1 | checkParameterDefaultValues | 0.0013–0.002 | 0.0016–0.0025 |
| distinct | 1 | checkNonConstantParameterDefaults | 0.0007–0.0008 | 0.001–0.0017 |
| distinct | 1 | complete-module-diagnostics | 0.2249–0.2254 | 0.2555–0.2784 |
| distinct | 100 | checkParameterDefaultValues | 0.3235–0.3285 | 0.136–0.1362 |
| distinct | 100 | checkNonConstantParameterDefaults | 0.2671–0.2716 | 0.0699–0.0721 |
| distinct | 100 | complete-module-diagnostics | 3.5683–3.632 | 3.0913–3.1779 |
| distinct | 1000 | checkParameterDefaultValues | 19.4927–20.0342 | 1.4224–1.4335 |
| distinct | 1000 | checkNonConstantParameterDefaults | 19.4203–20.4518 | 0.7332–0.7498 |
| distinct | 1000 | complete-module-diagnostics | 120.691–121.3883 | 73.3115–73.4683 |

At 1,000 procedures, each rule's project-name reads drop from 1,002,000 to 1,002 for repeated types and 2,001 for distinct types. The latter includes per-distinct-type result display names. The repeated one-procedure nonconstant check is slower, as are the distinct one-procedure nonconstant and complete calls; distinct one-procedure compatibility ranges overlap. Repeated and distinct one-procedure fixtures have identical source shapes but occur at different points in the warmed process, so they are not independent cold comparisons. No editor, cold-start or heap improvement is claimed.

## Validation

- Types and 145 focused tests pass.
- Twenty-seven new tests: twelve work-bound tests fail on the old rule; fifteen semantic controls already pass. Tests cover both rules across repeated/distinct project types, exact diagnostics/spans, supported object kinds, excluded kinds, duplicate ambiguity, host priority, unknown qualified types, unused metadata, inactive procedures and metadata replacement between queries.
- Full suite: 15,359 tests across 785 files pass; 33 tests and seven files skipped.
- 40,000 generated complete public-rule results, 2,000 generated complete-module results and 24,768 oracle module diagnostic/error results match the pinned baseline, with no differences. Generated cases vary declarations, defaults, arrays, absent defaults, eligible/ineligible duplicate metadata and active/inactive/unknown directives. Oracle comparison covers 8,256 sources across LF/CRLF/CR.
