# Directive validation and compiler-constant replay

The directive audit found two independent sources of quadratic work:

- Directive trailing-statement validation walked from token zero for every directive, even though the relevant tokens begin at the directive end. Use the existing firstTokenAtOrAfter binary lookup and scan only that physical line.
- Compiler-constant evaluation copied all preceding module/project constants into a new map for every expression. Constant indexing additionally converted the accumulated map to an object and back. The expression parser only needs lookup, so query the current project/module table first and the compiler table second, without copying the growing table.

The lookup preserves case folding, module/project precedence over compiler constants, falsy/special values, and forward definition order. Constant indexing still treats a missing name as zero; activity replay keeps names unknown when the caller has not supplied project constants. There is no additional persistent cache, and caller environments remain unchanged.

## Reproduction

Run node scripts/benchmark-directive-forms.mjs --rounds=15. The runner bundles local source with the existing esbuild dependency, parses fixtures before timing, warms each operation three times and reports median/p95 over 15 samples. Compare with commit f7aa0bb3 using the same runner. No test suite ran concurrently with either benchmark.

Node 24.18.0, AMD Ryzen 7 9800X3D, Windows. Times are milliseconds:

| Fixture | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: |
| 100-const-directives | 0.421 | 0.238 | 0.749 | 0.687 |
| 1000-const-directives | 12.189 | 1.173 | 14.34 | 1.542 |
| 5000-const-directives | 387.793 | 6.067 | 412.141 | 7.722 |
| 1000-trailing-statements | 15.527 | 1.709 | 20.783 | 2.341 |
| 1000-dependent-constant-index | 63.177 | 1.478 | 66.795 | 1.945 |
| 1000-dependent-constant-tracker | 11.893 | 1.456 | 13.099 | 2.589 |

The 5,000-directive rule pass is approximately 64 times faster; indexing 1,000 dependent constants is approximately 43 times faster. These are warm synthetic stage measurements that exclude initial parsing, not whole-workbook latency guarantees. The runner also checks diagnostic counts, final constant values and branch activity.

## Validation

Both complexity regressions fail against the original implementation: 751,500 token-array reads for 500 directives (budget 15,000), and 160,000 copied map entries for 400 dependent constants (budget 4,000). They pass after the fix. Semantic tests cover CR/LF/CRLF boundaries, comment-only colons, no final newline, source spans, dependent constants, missing-name policy, falsy/special-value shadowing and caller environment isolation.
Type checking passed. The full Vitest suite passed: 559 files, 11,993 tests passed and 13 skipped.
