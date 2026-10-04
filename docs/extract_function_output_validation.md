# Extract Method Function output declarations

Issue #893: a single write-first output becomes a Function result, but its assignment and return line referred to an undeclared helper local. The original scalar fixtures have no undeclared-variable diagnostic; the generated helper does.

Declare that output in the helper, preserving the original full local declaration clause. For a caller parameter, generate a local declaration from its name/type. An output declaration already inside the selection stays in the copied body. When the helper name equals the output name, its implicit Function result variable is sufficient. Multiple outputs remain ByRef parameters in a Sub.

Validation: 27 regression failures and six passing controls on baseline d2cc5892; all 33 now pass across LF/CRLF/CR. Cases cover scalar types, untyped Variant, fixed-length String, type suffixes, caller ByVal/ByRef parameters, case-insensitive implicit result names and multiple ByRef outputs. The real analyzer verifies undeclared-variable diagnostics; parsed ASTs verify one explicit declaration per independent helper output. 95 focused tests and type checking pass.

A frozen-AST native check covers 600 cases (200 variants across three line endings), ten scalar type forms, local and parameter outputs, input/multiple-output controls and implicit result names. It repairs 429 formerly undeclared outputs and preserves 171 complete control results and applied texts. It does not execute VBA in the VBE or claim that all Extract Method semantics are correct.

Full suite: 644 files passed, 13,087 tests passed, 13 skipped. The separate grouped-declaration repair is #886 / PR #892.

Performance guard: actual Extract Method, Node v24.18.0, AMD Ryzen 7 9800X3D 8-Core Processor           , baseline d2cc5892, 15 rounds after three warmups. All 16 complete benchmark result hashes match. These existing local-order fixtures contain inputs and moved scratch locals rather than single output Functions; they check unaffected work for regressions. Warm reuses parsed source; fresh adds a trailing source-key comment. Assertions run outside the clock. Timing is mixed and no speedup is claimed. Synthetic large parameter lists are not checked against VBE parameter limits.

| Case | Before median ms | After median ms | Before p95 ms | After p95 ms |
| --- | ---: | ---: | ---: | ---: |
| 1/0/ordered/read/warm | 0.02 | 0.022 | 0.043 | 0.043 |
| 1/0/ordered/read/fresh | 0.076 | 0.072 | 0.357 | 0.147 |
| 5/0/shuffled/read/warm | 0.018 | 0.018 | 0.057 | 0.035 |
| 5/0/shuffled/read/fresh | 0.076 | 0.072 | 0.314 | 0.298 |
| 100/0/shuffled/read/warm | 0.216 | 0.225 | 0.381 | 0.361 |
| 100/0/shuffled/read/fresh | 0.835 | 0.741 | 1.244 | 1.241 |
| 1000/0/ordered/read/warm | 3.701 | 3.653 | 4.262 | 4.674 |
| 1000/0/ordered/read/fresh | 7.593 | 7.646 | 11.223 | 9.351 |
| 1000/8000/shuffled/read/warm | 3.953 | 3.887 | 5.882 | 5.418 |
| 1000/8000/shuffled/read/fresh | 6.33 | 5.983 | 8.914 | 7.873 |
| 1000/100000/shuffled/read/warm | 4.796 | 4.771 | 6.965 | 6.223 |
| 1000/100000/shuffled/read/fresh | 8.439 | 8.262 | 15.046 | 10.985 |
| 3000/100000/shuffled/write/warm | 27.458 | 27.777 | 30.389 | 28.688 |
| 3000/100000/shuffled/write/fresh | 36.19 | 36.044 | 47.114 | 43.694 |
| 1000/100000/shuffled/few/warm | 0.478 | 0.496 | 0.539 | 1.729 |
| 1000/100000/shuffled/few/fresh | 2.869 | 2.808 | 4.395 | 5.1 |

Reproduce sequentially:

```powershell
node scripts/benchmark-extract-method.mjs --local-order --baseline=d2cc5892
node scripts/benchmark-extract-method.mjs --local-order
```
