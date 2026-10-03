# Extract Method reference scanning performance

Extract Method previously scanned the stripped module separately for every declared local and parameter. Each local used in the selection also triggered a separate full token-stream pass to classify its reads and writes. Large procedures therefore multiplied source scanning by their number of locals.

The fix collects occurrences for all candidate names in one source sweep, excludes declarations and references outside the containing procedure as before, and classifies all selected names together. The shared scanner preserves case-insensitive Unicode names, offsets, comment/string stripping and fresh returned arrays. Single-name consumers retain a direct string comparison in the hot loop. No additional persistent cache is introduced.

## Reproduction

Run node scripts/benchmark-extract-method.mjs --rounds=15 in the target checkout. The runner bundles the local source with the existing esbuild dependency, takes three warmups and reports the median and p95 of 15 samples. Each successful extraction touches five locals and leaves later reads in the original procedure. The baseline uses the same runner with the original implementations from commit b1603701.

Node 24.18.0, AMD Ryzen 7 9800X3D, Windows; times in milliseconds:

| Fixture | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: |
| 100 locals, 1,000 statements | 15.497 | 0.382 | 19.991 | 0.635 |
| 300 locals, 3,000 statements | 140.925 | 0.963 | 153.700 | 1.518 |
| 1,000 locals, 2,000 statements | 399.425 | 1.009 | 408.865 | 1.534 |
| Single-name scan, 10,000 statements (control) | 0.550 | 0.572 | 0.862 | 0.901 |

These are warm synthetic measurements, not a claim about every workbook. The largest fixture is approximately 396 times faster. The single-name control stays approximately unchanged; no speedup is claimed for it.

## Validation

The new regression test requires one reference-classification pass when extracting two locals from a procedure with 100 declarations. It fails against the original Extract Method implementation (two passes) and passes with the fix. Additional tests cover mixed line endings, Unicode/case variants, whole-name boundaries, bracketed names, strings, continued comments, empty queries, and independently owned result arrays. Existing extraction tests verify emitted edits and refusal behavior.

Type checking passed. The complete Vitest suite passed: 556 files, 11,980 tests passed and 13 skipped.
