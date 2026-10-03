# Attribute annotation rewrite audit

## Finding and fix

The rewriter scanned module lines again for every annotation target and inserted each attribute by shifting the line array. Modules with many annotated members paid quadratic target-search costs, including unchanged saves.

ModuleText now holds linked line records and lazily indexes procedure occurrences and declaration-section variables. References stay valid as attributes are inserted. Continued-header endpoints and owned attribute runs are reused; duplicate attributes retain their first-match behavior. Attribute edits still execute in annotation order, and new member attributes are prepended below the target header as before. The module header lookup retains its original matching semantics.

The representation adds transient line records and target indexes. It is local to a rewrite and is not a persistent cache. The production save path already skips applyAttributeAnnotations when there are no recognized annotations, so ordinary saves without annotations do not allocate this representation.

## Measurements

Run node scripts/benchmark-attribute-rewriting.mjs --rounds=15 in the target checkout. Baseline: 0c39d113. Node v24.18.0 on AMD Ryzen 7 9800X3D; three warmups and 15 measured samples, original and optimized runs sequentially, without concurrent tests. Fixtures have two annotations per procedure and ten body statements. Annotation reading is outside the timed public applyAttributeAnnotations call; line construction, edits, and output joining are included.

| Fixture | Before median / p95 ms | After median / p95 ms |
| --- | ---: | ---: |
| 100 procedures, insert | 2.045 / 2.456 | 0.128 / 0.476 |
| 100 procedures, unchanged | 1.987 / 2.547 | 0.194 / 0.338 |
| 1,000 procedures, insert | 199.582 / 222.052 | 1.529 / 1.769 |
| 1,000 procedures, unchanged | 182.760 / 196.450 | 1.856 / 2.404 |
| 3,000 procedures, insert | 1749.037 / 1883.669 | 4.594 / 8.999 |
| 3,000 procedures, unchanged | 1655.327 / 1676.026 | 5.900 / 8.946 |

The 3,000-procedure insertion stress fixture improves about 381x. These measurements isolate annotation rewriting; container serialization, filesystem writes, annotation reading, and whole-editor save latency are not measured. The new indexes make target lookup proportional to module size plus queries; arbitrarily many module-header annotations and other processing outside these indexes are not covered by a general linear-time claim.

## Validation

The original implementation fails the procedure-header work bound with 159,000 reads for a 100-procedure fixture. The optimized implementation passes. Existing annotation and property-leg tests pass, as do new LF/CRLF tests for repeated edits, variable scope, insertion order, and unrelated owned attributes. An additional deterministic comparison of 500 generated modules matches exact output text, change records, and skipped-target messages against the original implementation, including duplicate procedure names, continued headers, duplicate attributes, standard/class headers, missing targets, and repeated edits.

TypeScript check and the complete suite pass: 569 files, 12,086 tests, 13 skipped.
