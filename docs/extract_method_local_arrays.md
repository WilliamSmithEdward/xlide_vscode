# Extract Method local array bindings

Baseline: `3601ffc7267614130f5a2a968c222ac613a25467` (published PR #1223).
The published baseline tree equals the tested parent head
`704eb8b11c8a4c555b518121ff7e50104acdc48f`.

Extract Method used the element type of a local array as though it were the
variable's complete type. Extracting `Debug.Print x(1)` for `Dim x(1 To 2) As
Long` produced `Private Sub Work(ByVal x As Long)`. Read/write arrays could
produce a scalar ByRef parameter. A write-first dynamic array read afterward
could produce `Private Function Work() As Long` and `Work = x`, attempting to
return the whole array as a scalar. These generated forms are invalid VBA.

Touched supported local arrays now stay declared in their original invocation
and are passed through a typed ByRef array parameter, such as `ByRef x() As
Long`. The original declaration retains fixed bounds or dynamic allocation,
while the helper receives the array binding. Element writes, ReDim Preserve,
and Erase continue to affect the original array. Selected array declarations
remain in the caller rather than being duplicated in the helper. Arrays are
excluded from scalar input/output and declaration-movement heuristics.

The existing parameter-declaration helper now also handles local array
bindings, preserving type suffixes, qualified/bracketed spelling, and omitted
As clauses under module DefType. Call arguments retain their declared name
spelling. Scalar Function results can coexist with array helper parameters.

As New arrays and fixed-length String arrays are explicitly refused when
touched. Their ordinary helper parameter forms did not preserve the original
binding in native checks: the original source completed, while each extracted
candidate stopped in the VBE. Neither candidate was retained. The As New
formal loses automatic object creation; a variable-length String-array formal
cannot express the original fixed-length element declaration. An unrelated
scalar selection still works in procedures containing these declarations.
Existing Static and touched-ParamArray restrictions remain.

## Validation

- Type checking, 298 focused Extract Method tests, and all 16,100 tests across
  812 files pass (33 tests / seven files skipped).
- 60 dedicated cases cover LF, CRLF, and CR. On the baseline, 54 fail and six
  unrelated-scalar controls pass. All 60 pass with the repair.
- Coverage includes read-only/read-write/write-first arrays, fixed/dynamic and
  multidimensional bounds, ReDim/Preserve, Erase, object arrays, declaration
  suffixes, DefType, bracketed/qualified declarations, selected Dim statements,
  sibling declarations, a scalar Function result, multiple array/ordinary
  parameter bindings, and unsupported-declaration refusals. Bracketed editor
  fixtures test source preservation; they are not native-grammar claims.
- Three existing array-movement expectations now assert that the array stays
  in the caller and uses a ByRef array helper, while its movable object sibling
  still moves. Scalar declaration-movement tests remain unchanged.
- Complete public results and applied text are unchanged for 500 generated
  scalar-only modules and 24,768 corpus-context scalar extractions, spanning
  8,256 source environments under three line endings. This comparison does
  not assert equality for intentionally changed array extraction.
- Previously validated parameter sources also remain byte-identical to the
  sources from the prior repair's 32 native runs.
- 24 native Excel executions ran the actual original and generated sources in
  a disposable workbook: 12 cases, each before and after extraction. Every
  pair matched an independently specified expected result:

| Case | Original and extracted result |
| --- | --- |
| Fixed array read | `3` |
| Fixed array read/write | `4` |
| Dynamic array allocation and element write | `1:7` (lower bound and value) |
| ReDim Preserve | `7:9` (retained and added element) |
| Erase fixed array | `0` |
| Multidimensional array | `3:7` (second lower bound and value) |
| Variant array | `hello` |
| Integer declaration suffix | `7:2` (value and VarType) |
| Untyped array under DefLng | `7:3` (value and VarType) |
| Collection array | `7` |
| Array Dim inside the selection | `7` |
| Array mutation plus scalar Function output | `7:8` |

Known invalid scalarized baseline output was inspected through the public
refactoring API rather than executed. A bracketed local-name fixture that
stopped in the VBE even before extraction was excluded from native evidence.

This is a correctness repair found during the analyzer audit. No timing,
memory, editor-latency, or complete analyzer/repository coverage claim is made.
