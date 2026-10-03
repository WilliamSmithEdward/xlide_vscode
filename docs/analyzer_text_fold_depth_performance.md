# Recursive value folding recovery

Issue: #820. Baseline: `67dc5c1e4256adbcc96e551eb42b36a612d97f31`.

Assignment text folding recursively visits conversion arguments without a limit. The shared condition-value evaluator also recursively creates new parsers for call arguments, and directly recurses for parentheses, Not and unary minus. Deep CStr input can therefore fail in many diagnostic rules before reaching the assignment text fold; StrConv reproduces the independent text-fold failure.

Both evaluators now use MAX_EXPRESSION_DEPTH (256). Each argument or recursive parenthesized/prefix expression consumes depth. Independent siblings retain their parent's depth. Beyond the budget a value is unknown, and later statements remain analyzable. This intentionally stops folding expressions that previously happened to work above the limit; it preserves ordinary folds within it. The generated 5,000-call inputs exceed VBA's physical-line limit and test malformed editor input, not runtime-valid VBA.

Run `node scripts/benchmark-text-fold-depth.mjs`, then the same command with `--baseline=67dc5c1e4256adbcc96e551eb42b36a612d97f31`. The temporary bundle exposes the private text fold solely to measure it; production exports are unchanged. Tokens are prepared outside timing. Text-fold measurements exclude the rest of the analyzer; Abs measures numberValue. Each case uses three warmups and fifteen timed samples on Node 24.18.0, Ryzen 7 9800X3D. No audit tests or other audit benchmarks ran during measurement.

| Depth/function | Baseline median / p95 ms | Fixed median / p95 ms |
| --- | ---: | ---: |
| 10 CStr | 0.021 / 0.036 | 0.024 / 0.034 |
| 10 StrConv | 0.020 / 0.028 | 0.021 / 0.025 |
| 10 Abs | 0.018 / 0.038 | 0.017 / 0.043 |
| 100 CStr | 0.405 / 2.142 | 0.407 / 2.213 |
| 100 StrConv | 0.297 / 0.478 | 0.287 / 0.499 |
| 100 Abs | 0.336 / 0.673 | 0.318 / 0.437 |
| 255 CStr | 1.618 / 3.183 | 1.515 / 2.354 |
| 255 StrConv | 2.000 / 2.983 | 1.894 / 3.413 |
| 255 Abs | 1.857 / 2.141 | 2.003 / 2.303 |
| 5,000 CStr | 493.750 / 533.538 | 62.468 / 76.113 |
| 5,000 StrConv | 598.015 / 669.106 | 110.741 / 153.289 |
| 5,000 Abs | 166.138 / 200.281 | 74.195 / 82.758 |

All baseline 5,000-call samples overflowed the stack; all fixed samples returned unknown. Ordinary controls have mixed small timing differences, including a slower 255-level Abs control. This is a recovery improvement, not a claim of a universal speedup or constant-time handling of enormous token input: token splitting still scans the remaining arguments.

Validation covers both shared-limit boundaries, mixed conversion calls, concatenation siblings, deeply nested CStr/StrConv full analysis with a later diagnostic, numeric and IIf argument nesting, parentheses, prefix operators and a mixed call/parenthesis budget. A generated differential compared 1,000 complete diagnostic outputs below the budget, including literals, known locals, date/name stand-ins, source shadowing, substring/case folds and ordinary conditions; outputs matched exactly. Type checks and 177 focused tests passed.

The full suite passed: 610 test files, 12,705 tests passed and 13 skipped. The pathological full-analysis regressions have a 60-second timeout for parallel suite load; the initial 20-second timeout was too short despite passing in isolation. `git diff --check` passed, and the branch merges cleanly with main `041fa29f`.
