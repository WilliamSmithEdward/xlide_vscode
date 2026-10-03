# Long numeric logical-expression audit

## Finding and fix

The overflow rule recursively split logical expressions at the last lowest-precedence operator. Every split rescanned and copied the preceding token prefix. This made flat chains quadratic despite the earlier one-pass operator selection, and a 5,000-operand chain could overflow the JavaScript stack.

The fix collects top-level logical operators once and folds disjoint operands from left to right, reducing higher or equal precedence before the next operator. Logical folding no longer recursively descends through flat operator chains. Left associativity, numeric conversion and result types, eager operand evaluation, first unknown/overflow propagation, and the source span of each reduction are preserved. Arithmetic-only expressions avoid operator and value-stack allocations.

Structural nesting still follows the parser recovery limit. The claim here is linear handling of top-level logical operators and disjoint operand copying; folding arbitrary nested operands and other analyzer rules is outside that claim.

## Measurements

Run node scripts/benchmark-overflow-folding.mjs --rounds=15. Baseline: 8a319d35, including the depth-recovery and intrinsic-argument fixes. Node v24.18.0 on AMD Ryzen 7 9800X3D; three warmups and 15 samples, no concurrent test process. Timed public checkOverflow rule passes use parsed modules and prebuilt symbols, with warm analyzer caches.

| Fixture | Before median / p95 ms | After median / p95 ms |
| --- | ---: | ---: |
| 100-operand And chain | 0.534 / 2.199 | 0.221 / 1.024 |
| 1,000-operand And chain | 22.783 / 26.697 | 1.226 / 1.877 |
| 3,000-operand And chain | 212.861 / 224.630 | 2.581 / 4.052 |
| 300 procedures, arithmetic | 9.937 / 11.900 | 9.647 / 12.176 |
| 10,000 arithmetic assignments | 26.048 / 28.406 | 25.965 / 28.454 |
| 1,000 assignments, 100 arithmetic operands | 32.349 / 36.975 | 33.032 / 34.251 |
| 1,000 mixed logical assignments | 4.972 / 7.451 | 4.760 / 5.519 |

The largest chain improves about 82x in this synthetic stress fixture. The unbroken long physical lines exceed VBA source-line limits; they exercise editor/recovery input and isolate folder scaling. These measurements exclude parsing and do not represent end-to-end analyzer latency.

## Validation

The original implementation fails the token-copy bound (639,200 entries copied for a 400-operand chain) and throws on the 5,000-operand recovery case. The optimized implementation passes both and retains the later Const overflow diagnostic. Additional tests preserve first-failure spans and unknown propagation, including LongLong operands whose right-hand expression overflows first. Existing precedence, numeric overflow, intrinsic argument, and structural recovery suites pass. A deterministic comparison of 1,000 generated expressions matches original diagnostic kinds, messages, and spans exactly.

Full repository suite: 569 files, 12,089 tests passed, 13 skipped. TypeScript check passed.
