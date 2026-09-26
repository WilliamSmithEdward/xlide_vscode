# VBA Runtime Type Mismatch Oracle Matrix

Date: 2026-05-31

Purpose: isolate when VBA accepts source at compile time but raises deterministic
runtime error 13 (`Type mismatch`) for literal scalar coercion. These cases are
evidence for red XLIDE diagnostics whose `vbeCompileEquivalent` value is false.

## Hypothesis

If XLIDE can prove a string literal cannot be coerced to the required numeric or
Boolean scalar type, the code may compile but will raise runtime error 13 when
executed. Numeric strings, Boolean strings, stringification, and `&`
concatenation should remain accepted controls. String-plus-string with `+`
should also remain accepted if both operands are string values.

## Method

- Run focused `mode: "run"` Excel/VBE oracle fixtures only.
- Detect modal `Microsoft Visual Basic` runtime dialogs as the primary oracle
  signal.
- Record each fixture with `evidencePhase: "runtime"` and
  `diagnosticMeaning: "runtime-error"` or `"runtime-valid"`.
- Keep compile acceptance recorded separately with `evidencePhase: "compile"`.

## Results

| Fixture | Variable Isolated | Outcome |
| --- | --- | --- |
| `integer_assignment_nonnumeric_string_runtime` | Direct Integer assignment from `"string"` | Rejected: runtime error 13 |
| `nonnumeric_string_to_double_assignment_runtime` | Direct Double assignment from `"blah"` | Rejected: runtime error 13 |
| `boolean_assignment_nonboolean_string_runtime` | Direct Boolean assignment from `"maybe"` | Rejected: runtime error 13 |
| `nonnumeric_string_to_currency_argument_runtime` | Same-module Currency argument from `"blah"` | Rejected: runtime error 13 |
| `left_length_nonnumeric_string_runtime` | Native `Left` Length argument from `"bad"` | Rejected: runtime error 13 |
| `numeric_plus_nonnumeric_string_assignment_runtime` | Right operand `"string"` in `1 + "string"` | Rejected: runtime error 13 |
| `numeric_plus_nonnumeric_string_left_runtime` | Left operand `"string"` in `"string" + 1` | Rejected: runtime error 13 |
| `numeric_plus_nonnumeric_string_parenthesized_runtime` | Parenthesized `1 + "string"` | Rejected: runtime error 13 |
| `numeric_multiply_nonnumeric_string_runtime` | Operator changed from `+` to `*` | Rejected: runtime error 13 |
| `integer_assignment_numeric_string_runtime` | Direct Integer assignment from `"2"` | Accepted |
| `double_assignment_numeric_string_runtime` | Direct Double assignment from `"2.5"` | Accepted |
| `numeric_string_to_currency_argument_runtime` | Same-module Currency argument from `"100"` | Accepted |
| `left_length_numeric_string_runtime` | Native `Left` Length argument from `"2"` | Accepted |
| `boolean_assignment_true_string_runtime` | Direct Boolean assignment from `"True"` | Accepted |
| `number_to_string_assignment_runtime` | String assignment from numeric literal `123` | Accepted |
| `numeric_plus_numeric_string_runtime` | Numeric arithmetic with `"2"` | Accepted |
| `string_concat_nonnumeric_string_runtime` | `&` concatenation with `"string"` | Accepted |
| `string_plus_string_literals_runtime` | `+` between two string literals assigned to `String` | Accepted |

## Conclusion

The evidence supports red deterministic-runtime diagnostics for nonnumeric string
literals in numeric/Boolean assignment, numeric/Boolean argument coercion, and
numeric arithmetic expressions. The rule remains narrow: numeric strings,
Boolean strings, ordinary stringification, string-plus-string with `+`, unknown
values, `Variant`, and host-dependent coercions are not hard errors.
