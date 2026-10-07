# Extract Method procedure parameter bindings

Baseline: `5e1769b068d7534fac8d2c91fbe7891c32926c34` (published PR #1221).

Extract Method previously classified procedure parameters like local variables:
a value written in the selection was considered an output only if the original
procedure read it afterward. A caller can observe a ByRef parameter even when
there are no later reads inside that procedure. For example:

```vb
Option Explicit
Public Sub Main(ByRef x As Long)
x = x + 1
End Sub
```

The old helper accepted `ByVal x As Long`, losing the caller's update. A
write-first selection (`x = 1`) instead produced a parameterless helper with an
undeclared `x`. Both explicit and default ByRef parameters were affected.

The helper now accepts every touched ordinary procedure parameter ByRef. It
receives the variable that already belongs to the original invocation. An
original ByVal parameter is already a private copy, so passing that variable
ByRef to the helper does not change the original caller's variable. Passing
aliases directly also preserves intermediate observations and writes made by
other procedures, without relying on local read/write classification.

Original array shape, declaration suffix, bracketed name/type, and explicit
qualified type are retained. Untyped parameters keep their original spelling in
the same module, preserving DefType inference. Optional markers/defaults are
not copied: the original procedure has already applied its default or missing
argument value. Native checks verified that forwarding a missing Optional
Variant retains `IsMissing` behavior.

Selections touching ParamArray parameters are refused with an explanation.
Native attempts to forward the original ParamArray directly to either a typed
Variant-array parameter or a ByRef Variant helper did not complete without a
VBE interruption; neither candidate was retained. This change does not claim to
implement general ParamArray forwarding. A selection that does not touch that
parameter is unaffected.

Local input/output classification remains unchanged, including a single local
Function result alongside touched procedure parameters. Local arrays and other
unrelated extraction restrictions are outside this repair's validated scope.

## Validation

- Type checking, 238 focused Extract Method tests, and all 16,040 tests across
  811 files pass (33 tests / seven files skipped).
- 63 dedicated parameter cases fail on the baseline and pass with the repair,
  covering LF, CRLF, and CR, write-first/read-modify-write/read-only parameters,
  aliases, indirect writes, arrays, Optional declarations, suffixes, bracketed
  and qualified declarations, continued headers, DefType, ParamArray refusals,
  and a local Function result alongside a parameter update.
- The six obsolete parameter-as-Function-output expectations were replaced by
  the parameter tests. Local Function-output coverage remains in place.
- Complete public results and applied text are identical to baseline for 500
  generated local-only modules and 24,768 corpus-context local extractions
  (8,256 source environments under three line endings). These comparisons do
  not assert equality for the intentionally changed parameter behavior.
- 32 native Excel runs executed the actual original and generated sources in
  a disposable workbook: 16 cases, each before and after extraction. Every
  pair matched its independently specified result:

| Case | Result before and after |
| --- | --- |
| ByVal Long, write-first and increment | `0` (caller's value retained) |
| Explicit/default ByRef Long, write-first and increment | `1` |
| Two parameters aliasing one Long | `2` |
| Another procedure mutating the parameter | `7` |
| Dynamic Long array resized and written | `1:7` (lower bound and element) |
| Optional Variant, omitted then provided | `True:False` (`IsMissing`) |
| Optional ByVal default and caller isolation | `6:1:0` |
| Integer declaration suffix | `7` |
| Untyped parameter under DefLng | `7:3` (value and VarType) |
| ByVal Collection rebinding | `3` (original caller object) |
| ByRef Collection rebinding | `7` (replacement object) |

The baseline native probe had independently demonstrated the defect: explicit
and default ByRef increments returned `0` after extraction versus `1` before;
the ByVal control returned `0` in both. Known invalid baseline write-first
output was inspected as generated source rather than executed.

This is a correctness repair found during the analyzer audit. No performance,
editor latency, cold-start, memory, or whole-repository coverage claim is made.
