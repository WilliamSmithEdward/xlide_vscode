# Numeric folding recovery and intrinsic argument reuse

## Confirmed findings

Numeric folding could overflow the JavaScript stack on 1,000 nested parentheses or 1,500 nested CInt/Sgn calls. The analyzer's per-rule guard catches the exception, but later Const and assignment overflow findings are lost. Numeric folding and nested-part discovery now share the expression parser's 256-level recovery limit. Beyond that limit, the expression stays unknown and analysis continues. This intentionally limits recovery input; it does not claim to evaluate arbitrary nesting.

Intrinsic function result folding also evaluated each argument twice: once to check every argument for overflow, then again to obtain the selected value. Ten nested Sgn calls therefore multiply repeated evaluations. The argument results are now retained only for the current intrinsic invocation, including unknowns. All arguments are still eagerly checked, preserving overflow reports from unselected IIf and Choose arguments.

## Measurements

Run node scripts/benchmark-overflow-folding.mjs --rounds=15 in each target checkout. Baseline: e03bba11, which includes the earlier logical-scan improvement. Node v24.18.0 on AMD Ryzen 7 9800X3D; three warmups and 15 samples, no concurrent test process. Parsing and symbol construction are excluded; the public overflow rule is timed with warm analyzer caches.

| Fixture | Before median / p95 ms | After median / p95 ms |
| --- | ---: | ---: |
| 100 Const declarations with 10 nested Sgn calls each | 120.568 / 126.910 | 1.801 / 2.476 |
| 300 procedures with arithmetic assignments | 9.553 / 10.140 | 10.188 / 18.942 |
| 10,000 short arithmetic assignments | 26.387 / 27.629 | 26.378 / 40.086 |
| 1,000 assignments with 100 operands | 31.541 / 35.744 | 32.383 / 34.375 |
| 1,000 mixed logical assignments | 4.758 / 6.538 | 5.065 / 6.297 |

The nested intrinsic fixture improves about 67x. Ordinary controls range from effectively unchanged to about 7% slower in these samples, and p95 timings are variable; no broad speedup claim is made. These are isolated-rule measurements, not editor or full-analyzer latency.

## Validation and remaining audit

Recovery regressions retain later Const and assignment overflow findings after excessive parentheses, conversions, intrinsics, and Not chains; a procedure case checks deeply nested unknown calls. Ordinary nested overflow spans remain unchanged. Existing numeric folding, intrinsic argument, logical precedence, and parser-depth tests pass. A comparison of 500 generated intrinsic, conversion, division, and logical expressions matches original diagnostic kinds, messages, and spans. The full suite passes: 567 files, 12,058 tests, 13 skipped.

Long flat logical chains remain a separate performance candidate: an exploratory 3,000-operand And chain took about 274 ms in the isolated rule. Its recursive logical splitting is not changed here. Deeper structural folding outside the supported recovery limit intentionally remains unknown.
