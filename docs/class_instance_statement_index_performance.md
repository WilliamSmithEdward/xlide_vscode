# Class-instance statement indexing

The class-instance rule collected Set statements separately for every local and scanned all statements again to determine whether each selected class instance escaped. With N locals and N statements, these walks were quadratic. Even scalar locals that could never produce a tracked class instance incurred the Set scan.

Normalize eligible declarations once, skip unsupported scalar types, and index Set targets and escaping uses in one query-owned pass. Preserve Set multiplicity, As New eligibility, field-assignment suppression, bracketed names, active branches and original diagnostic spans. Remove the unused private Instance.name field. No cross-query cache is introduced.

## Reproduction and measurements

Run `node scripts/benchmark-class-instance-index.mjs`, and repeat with `--baseline=a8fcd2dd6d3478c21d65a16deae9bd8f74eb3ce1`. The baseline overrides only this rule while retaining all other current dependencies. Each procedure has N local declarations and N `Debug.Print 1` statements. Class locals use `As New Class1`; scalar locals use `As Long`.

Sequential before/after/after/before runs on Node 24.18.0, Ryzen 7 9800X3D, with three warmups and nine measured calls per row. Values are ranges of the two run medians in milliseconds.

| Locals | Count | Scope | Before ms | After ms |
| --- | ---: | --- | ---: | ---: |
| scalar | 1 | rule | 0.0043–0.0048 | 0.0022–0.0023 |
| scalar | 1 | complete-module-diagnostics | 1.0758–1.2259 | 1.0691–1.0745 |
| scalar | 100 | rule | 0.1532–0.1602 | 0.0091–0.0125 |
| scalar | 100 | complete-module-diagnostics | 4.5766–4.7889 | 4.5935–4.7151 |
| scalar | 1000 | rule | 9.2543–10.4893 | 0.0621–0.0664 |
| scalar | 1000 | complete-module-diagnostics | 40.4789–41.7718 | 29.149–30.3893 |
| class | 1 | rule | 0.0015–0.0015 | 0.0023–0.003 |
| class | 1 | complete-module-diagnostics | 0.4673–0.4744 | 0.5295–0.5827 |
| class | 100 | rule | 0.3986–0.5886 | 0.0615–0.0628 |
| class | 100 | complete-module-diagnostics | 3.3222–3.5214 | 2.9272–2.9936 |
| class | 1000 | rule | 36.6472–36.7804 | 0.2815–0.2896 |
| class | 1000 | complete-module-diagnostics | 63.0415–63.7858 | 26.0135–26.1986 |

At 1,000 locals, the old rule visits 1,000,000 statement rows for scalar declarations and 2,000,000 for class declarations through its filter/every callbacks. Both repeated walks are gone. Independent regression tests bound total token-name/word queries by 8N+10, so replacing callback walks with equivalent nested loops cannot satisfy the test.

Small class cases are slower, and the scalar 100-local complete-call ranges overlap. This is not an overall editor latency claim. AST/symbol preparation is outside direct-rule timing; complete-module calls include analyzer preparation and independent output assertions. The harness independently expects no direct-rule diagnostics, exactly N unused-variable diagnostics with the correct names, and no internal errors, then compares full results each round. No cold-start or heap measurements are claimed.

## Validation

- Type checking and 67 focused tests pass on the final source.
- 20 new tests: six bounded-work cases fail on the previous implementation; 14 independent semantic controls pass on both.
- Full suite: 15,288 tests pass across 780 files, with 33 tests and seven files skipped. This run precedes only removal of the unused Instance.name field; final focused checks cover that cleanup.
- 20,000 generated complete rule results and 24,768 complete module diagnostic/error results from 8,256 oracle sources across LF/CRLF/CR match the previous implementation. Cases cover Set multiplicity, declaration types, scalar/passing escapes, bracketed names, branches and assignment suppression.
