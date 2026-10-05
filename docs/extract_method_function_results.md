# Extract Method Function and Property Get result bindings

Baseline: `a17b68cca12850487aef8c908108e70b07af10f2` (published PR #1225).
The published baseline tree equals its tested parent head.

A Function or Property Get returns through a variable with its own name. That
variable belongs to the invocation, rather than module scope. Extract Method
previously treated it as a free reference: extracting `Compute = 7` from a
Function emitted a parameterless Sub still assigning `Compute = 7`. That
assignment no longer denoted the original Function result.

There is also a silent, native-confirmed corruption when a Property Let with
the same name exists. The old helper's assignment invokes the setter instead
of writing the getter's result. The original getter returns 7 and leaves the
setter's stored state at 0 (`7:0`); actual baseline-generated code returns 0
and writes 7 through the setter (`0:7`). The repaired generated code retains
`7:0`.

The helper now accepts the original result variable ByRef under a separate,
unique name, such as `ComputeResult`. Only result-variable references in the
selected statements are rewritten. Recursive calls to Compute remain calls
to the original procedure; qualified member/type references, comments, and
strings retain their source spelling. Result objects used as dot or bang
member receivers retain their binding as well.

The declaration keeps the return type's raw qualified/bracketed spelling and
type suffix. Untyped result names keep the original first letter, preserving
selective module DefType inference. Generated names avoid source names and
the chosen helper name and use a bounded prefix without splitting Unicode
code points. Scalar local Function outputs can coexist with the original
result variable's ByRef parameter.

New name/reference scans operate on the selected statement text. Existing
reference-kind classification uses absolute source offsets; cached module
tokens supply receiver checks and collision avoidance. No symbol-model or
project loading is introduced.

Selections transferring a typed array-valued result are refused: whole-array
assignment to a Function result cannot be represented as ordinary array
parameter assignment. A result-transferring selection that also exits the
original Function/Property is refused because that exit cannot be moved into
a helper. This does not implement general extraction of nonlocal control
flow or claim complete support for every VBA result/type expression.

## Validation

- Type checking, 379 focused Extract Method tests, and all 16,181 tests across
  813 files pass (33 tests / seven files skipped).
- 81 dedicated cases cover LF, CRLF, and CR: 75 fail on baseline and six
  unaffected recursive/qualified-call controls already pass. All pass after
  the repair. They cover writes, reads, indirect mutation, recursion, object
  receivers, Property Get, suffixes, DefType, qualified/bracketed type source,
  continued headers, collisions, long identifiers, strings/comments, scalar
  Function outputs, and unsupported result/exit refusals. Editor type-name
  fixtures are source-preservation checks, not native grammar claims.
- Complete public results and applied texts are unchanged for 500 generated
  local-only Sub extractions, 1,000 Function selections without result-variable
  references, and 24,768 corpus-context Sub extractions (8,256 source
  environments under three line endings). This does not assert equality for
  the intentionally changed result-transfer behavior.
- 30 native Excel runs executed actual original/generated sources in a
  disposable workbook: 15 cases, each before and after extraction. Every pair
  matched an independently specified expected result:

| Case | Original and repaired result |
| --- | --- |
| Scalar write | `7` |
| Result read/write | `4` |
| Indirect ByRef result mutation | `7` |
| Recursive Function | `6` |
| Object result creation and use | `7` |
| Integer suffix | `7:2` (value and VarType) |
| Selective DefLng C-C | `7:3` (value and VarType) |
| Variant and explicit String results | `hello:8` and `hello` |
| Property Get | `7` |
| Original result plus scalar local Function output | `15` |
| Object result plus scalar local Function output | `7:8` |
| Existing object result, dot/bang receiver only | `7` in both cases |
| Property Get with a same-name setter | `7:0` (result and setter state) |

Two additional native runs independently established the baseline setter
corruption (`7:0` original versus `0:7` baseline-generated). Known invalid
Function baseline output was inspected as generated source rather than run.

This is a correctness repair found during the analyzer audit. No measured
performance, memory, editor-latency, or complete surface-coverage claim is made.
