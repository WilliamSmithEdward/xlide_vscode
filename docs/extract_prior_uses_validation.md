# Extract Method earlier caller uses

Issue #895: a write-first scratch local whose value was not read after the selection took its declaration into the helper even when the caller had read or assigned it before the selection. The retained earlier statement then became undeclared under Option Explicit.

Record whether each touched local has an occurrence before the selected block, using the existing batched occurrence analysis. Such a clause remains declared in the caller; the helper still gets its independent declaration for the write-first value. The existing grouped-declaration edit builder retains these clauses alongside untouched siblings. Extraction remains available for valid earlier-use cases.

Validation: 30 cases across LF/CRLF/CR; baseline 24 failures and six passing controls, repaired all 30. Cases include earlier read/write/nested reads, first/middle/last grouped clauses, multiple retained clauses, untouched siblings, type suffix/fixed-length syntax, no-prior-use deletion and comment/string exclusions. The real analyzer verifies undeclared-variable diagnostics, and parsed ASTs verify declarations in both procedures. 134 focused tests and type checking pass.

Native frozen-AST verification covers 600 cases (200 variants across three line endings), 1–31 locals, groups/colon statements, Unicode identifiers, reads/writes, comments and strings. It repairs 480 formerly undeclared caller cases and preserves 120 complete control results/applied texts. Caller/helper declarations match expected clauses and every generated source edit list is non-overlapping. These are source/AST checks, not VBE execution or proof that all extraction semantics are correct.

Full suite: 645 files passed, 13,123 tests passed, 13 skipped. The baseline is merged grouped-declaration repair d2df8069 (PR #892); tree PRs #889 and #891 are integrated.

Performance guard: existing actual Extract Method local-order benchmark, Node v24.18.0, AMD Ryzen 7 9800X3D 8-Core Processor           , 15 rounds after three warmups. All 16 complete result hashes match. Warm reuses parsed source; fresh changes the source key with a trailing comment. Assertions run outside the clock. The extra occurrence check runs once per touched local; timings are mixed and no speedup is claimed. Large synthetic input signatures are not VBE parameter-limit checks.

| Case | Before median ms | After median ms | Before p95 ms | After p95 ms |
| --- | ---: | ---: | ---: | ---: |
| 1/0/ordered/read/warm | 0.026 | 0.026 | 0.061 | 0.048 |
| 1/0/ordered/read/fresh | 0.097 | 0.078 | 0.383 | 0.355 |
| 5/0/shuffled/read/warm | 0.022 | 0.019 | 0.041 | 0.042 |
| 5/0/shuffled/read/fresh | 0.082 | 0.073 | 0.448 | 0.342 |
| 100/0/shuffled/read/warm | 0.256 | 0.26 | 0.482 | 0.48 |
| 100/0/shuffled/read/fresh | 0.859 | 0.91 | 1.342 | 1.4 |
| 1000/0/ordered/read/warm | 3.807 | 3.811 | 4.581 | 4.567 |
| 1000/0/ordered/read/fresh | 8.062 | 7.91 | 9.878 | 10.565 |
| 1000/8000/shuffled/read/warm | 3.908 | 4.072 | 6.227 | 6.124 |
| 1000/8000/shuffled/read/fresh | 6.459 | 6.314 | 8.555 | 8.26 |
| 1000/100000/shuffled/read/warm | 4.764 | 4.816 | 7.442 | 7.151 |
| 1000/100000/shuffled/read/fresh | 8.425 | 10.463 | 10.764 | 12.438 |
| 3000/100000/shuffled/write/warm | 27.094 | 27.333 | 30.217 | 30.609 |
| 3000/100000/shuffled/write/fresh | 37.05 | 38.576 | 43.528 | 47.619 |
| 1000/100000/shuffled/few/warm | 0.509 | 0.495 | 1.729 | 2.119 |
| 1000/100000/shuffled/few/fresh | 2.781 | 2.892 | 4.488 | 4.596 |

Reproduce sequentially:

```powershell
node scripts/benchmark-extract-method.mjs --local-order --baseline=d2df8069
node scripts/benchmark-extract-method.mjs --local-order
```
