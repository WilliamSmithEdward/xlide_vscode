# Introduce Parameter with Optional and ParamArray signatures

Introduce Parameter previously appended a required `ByVal` parameter to every existing parameter list. This produced `required-param-after-optional` or `paramarray-not-last` diagnostics and could also pass the initializer into an old Optional parameter or the ParamArray tail.

For an existing Optional list, append an **Optional ByVal** parameter and pass its value explicitly by name at each known call. This skips old Optional parameters that the caller omitted and preserves the original parameter order. A safely convertible literal initializer also becomes the new default. Otherwise the default is the neutral literal `0` (or `""` for String); known callers still pass the actual initializer expression. All supplied project call-site edits must be applied together. Callers outside the supplied sources are outside this refactor's coverage, particularly for nonliteral initializers.

```vb
' Before
Public Sub Report(Optional ByVal mode As Long = 0)
    Dim limit As Long
    limit = 3
    Debug.Print limit
End Sub
Report 7

' After
Public Sub Report(Optional ByVal mode As Long = 0, Optional ByVal limit As Long = 3)
    Debug.Print limit
End Sub
Report 7, limit:=3
```

For a ParamArray list, insert the required parameter immediately before ParamArray and the initializer at the matching positional argument boundary. Keep all existing arguments, including omitted ParamArray entries, in place. If the call has no ParamArray entries, the initializer remains after the fixed arguments. If it has entries, only a literal known to fit the local type may move ahead of them; otherwise the refactor refuses with an evaluation/conversion-order explanation. This avoids generating invalid VBA or silently changing effects. This check is deliberately conservative about expressions, narrowing the automatic rewrite where safe conversion is unproved.

Named arguments cannot repair the ParamArray case: [Microsoft documents that all arguments to a ParamArray procedure must be positional](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/a-procedure-with-a-paramarray-argument-cannot-be-called-with-named-arguments). [Microsoft's parameter array documentation](https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/understanding-parameter-arrays) also requires ParamArray last. The repository's `docs/spec/[MS-VBAL].pdf`, §5.3.1.5, gives the Optional/ParamArray parameter-list grammar; §5.3.1.11 describes parameter-order binding and assignment.

The rewrite uses insertions, preserving nested calls, argument expression text, continuation trivia, and original ByRef parentheses. Existing binding filters, property-accessor refusals, recursion checks, and initializer scope checks remain in force. Call-site metadata now carries the bounded existing argument span; no lexer is added to the common call-site scan.

## Validation

Baseline: published main `3ebd180f45f486fec0aed810f1fde61642e3f620`, including concurrent parser-cache and tree changes.

The 49 dedicated tests cover all three line endings, empty/bare/Call/expression/named calls, fixed parameters, multiple Optional parameters and omissions, ParamArray omissions, nested calls, continuations, cross-module binding, safely convertible literal types, and refusal of effectful/unsafe conversions before existing ParamArray entries. The baseline fails 41 and passes eight controls. The corrected implementation passes all 49 and all 458 focused tests across 15 files.

Native validation executed the actual generated before/after source in a separate disposable workbook using the installed Excel instance. A `Mark` function records argument/initializer effects:

| Fixture | Before | After |
| --- | --- | --- |
| Optional arguments followed by effectful initializer | `MNL:1` | `MNL:1` |
| Two ParamArray entries and literal initializer | `MN:3:1` | `MN:3:1` |
| One fixed argument, empty ParamArray, effectful initializer | `ML:1:-1` | `ML:1:-1` |

These traces encode evaluation order, initialized value, and (for ParamArray) its upper bound. A required parameter inserted before the Optional parameters instead produces `LMN:1` even if the initializer's named argument is written last; the Optional tail avoids that reorder. Native validation covers these fixtures, not every possible host/type/expression.

Unchanged behavior was compared against the baseline on all 8,256 oracle source environments under LF/CRLF/CR. Complete public results, edited source, and cross-module edit lists matched for 24,768 Introduce Parameter calls with ordinary required signatures, and 24,768 companion Inline Variable calls. Zero differences. This is refactor regression coverage, not proof of complete analyzer or repository coverage. No performance or editor-latency claim is made for this correctness repair.

Final type checking passed. The integrated full suite passed 15,968 tests across 809 files (33 tests and seven files skipped), in 81.81 seconds. Regenerated native source fixtures were byte-identical after rebase and the final required-signature guard change.
