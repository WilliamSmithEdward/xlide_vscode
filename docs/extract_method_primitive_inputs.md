# Extract Method primitive input bindings

Tested parent: `6d2e2983bbec446f45f5c7a603ced3c4e2ab1669` (PR #1231).

The syntactic reference classifier intentionally labels a variable passed to a
procedure as a read: resolving its eventual ByRef/ByVal signature is outside
that classifier's contract. Extract Method treated this as proof that a primitive
input could safely be copied ByVal. Native checks demonstrate silent data loss:
a Long local returns 7 originally but 3 after extraction; a String local returns
seven originally but three afterward when a called ByRef procedure changes it.

Primitive inputs now retain their original variable through a ByRef helper
parameter, whether or not the selected code has a direct assignment. The existing
raw-declaration formatter preserves the formal type and escaped call-argument
spelling. The helper-binding formatter is shared across input/output groups. Original parameters, arrays,
Variant/object/opaque bindings, output selection, declaration placement and the
existing input-group ordering retain their prior handling. Primitive Function
outputs remain available alongside the preserved inputs.

A fixed-length String cannot be forwarded through an ordinary String formal
without changing its storage semantics. A native rejected prototype changed
`Len(x)` observed inside the selected block from 5 to 8 after a ByRef callee
assigned eight characters, even though the caller's final value was abcde in both
versions. The added variable-length formal delayed fixed-length copy-back.
Fixed-length Strings are therefore refused when extraction requires an input or
multiple-output helper parameter. Write-first fixed-length locals can still move
or become a single Function output: their helper-local declaration retains the
fixed length. An untouched declaration does not block another selection.

## Validation

- Type checking and 523 focused tests across 14 files pass.
- 51 dedicated tests cover LF, CRLF and CR. Of the initial 48 cases, the
  parent fails 45 and three fixed-length Function-output controls pass. Three
  additional bracketed-name cases test source preservation, not native grammar. Existing primitive input signature
  assertions now expect ByRef. The procedure-scan benchmark assertion also
  follows the corrected signature.
- 30 actual native Excel executions cover 15 original/generated pairs. Every
  pair matches an independent expected result:

| Case | Original and repaired result |
| --- | --- |
| Byte, Integer, Long, Single, Double, Currency | 7 with original VarType 17, 2, 3, 4, 5, 6 |
| Date | 7 with VarType 7 |
| Boolean | True with VarType 11 |
| String | seven with VarType 8 |
| ByRef call inside an expression | 7:1 (mutated input and result) |
| Mutation plus primitive Function output | 7:8 |
| Two indirectly mutated inputs | 7:8 |
| Mutation observed only inside selection | 7 |
| Fixed-length String Function output | abcde:5 |
| Primitive input beside an untouched fixed-length String | 7 |

- Four baseline native executions independently establish the Long/String losses.
  Two additional executions establish why the fixed-length ByRef prototype was
  rejected. These are distinct from the 30 successful repaired executions.
- 28,518 complete public-result/applied-text comparisons show zero differences
  for unchanged behavior: 500 generated write-first modules, 1,000 Function
  controls, 2,250 array/reference-capable procedure contexts, and 24,768 corpus
  queries (8,256 environments under three line endings). Intentionally changed
  primitive input signatures and fixed-length refusals are excluded.
- All 16,325 tests across 816 files pass (33 tests / seven files skipped),
  using a 15-second per-test timeout. The interrupted earlier run was not
  counted as validation.

All native work uses a disposable workbook. This correctness repair makes no
measured timing, memory, editor-latency or whole-analyzer/repository coverage
claim. Primitive variable identity for other output/movement paths and nonlocal
control flow remain audit candidates.
