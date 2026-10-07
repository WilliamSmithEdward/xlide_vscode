# Extract Method reference-capable local bindings

Baseline: `4dc6071b0b3a3b0964829ec8d124dd6725fb9f04` (published PR #1227).

Extracting a local Collection output produced a Function with `Work = x`
and a caller with `x = Work()`. These omit Set and cannot transfer an object
binding correctly. A Variant can hold either a scalar, object, or array, so
choosing a single Let/Set return mechanism is insufficient. Read-only-looking
inputs also lost indirect writes: native VBA checks of a Variant, a Collection,
and a Variant containing a Collection each returned 7 before extraction and 3
after the baseline extraction when a called procedure reassigned its ByRef
argument.

Touched Variant/untyped, object, and other nonprimitive locals now retain their
original variable in the caller and receive a ByRef helper parameter. The
existing raw declaration formatter preserves qualified/bracketed type spelling,
type suffixes, and omitted As clauses under DefType. Unknown types use their
explicit type; the repair does not assume an Enum or UDT is a class. These locals
are excluded from primitive scalar output and declaration-movement heuristics.
Primitive Function outputs can coexist with these bindings.

Touched As New declarations are refused: an ordinary formal parameter cannot
retain their automatic creation semantics. Same-named dot/bang members, including
implicit With members, are excluded from local binding classification. A helper
name colliding with a required parameter binding is also refused. Existing Static,
ParamArray, and fixed-length String-array restrictions remain.

## Validation

- Type checking and 81 dedicated tests pass, spanning LF, CRLF, and CR. The
  initial 66 cases yielded 60 failures and six passing controls on the baseline;
  15 additional cases guard qualified-member classification.
- 30 actual native Excel executions cover 15 original/generated pairs. Each
  matches an independent expected result: object outputs, Variant scalar/object/
  array values, untyped variables with and without DefType, indirect mutation,
  Enum and UDT locals, selected Dim, mixed object/primitive outputs, and object
  reads. The three indirect-write cases now preserve 7.
- Another 30 native executions rerun all 15 prior Function/Property Get result
  cases, including recursion, object dot/bang receivers and a same-named setter.
- Complete public results and applied text remain equal for 500 generated
  primitive-local Sub modules, 1,000 unchanged Function controls, and 24,768
  corpus-context queries (8,256 environments under three line endings). This
  comparison excludes intentionally changed bindings and collisions.
- All 16,259 tests across 814 files pass (33 tests / seven files skipped).
  An initial run
  passed 16,243 tests but hit the existing five-second timeout in the unrelated
  Buffer-accessor inventory test. The final invocation allows 15 seconds per
  test without changing repository test configuration.

The known invalid baseline object-return output was inspected through the public
API rather than executed. Native tests use a disposable workbook. Bracketed
editor fixtures establish source preservation, not native grammar acceptance.
This correctness repair makes no timing, memory, editor latency, or complete
analyzer/repository coverage claim. Primitive indirect writes and nonlocal control
flow remain separate audit candidates.
