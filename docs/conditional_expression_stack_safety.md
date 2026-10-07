# Conditional expression stack safety

`ConditionalExpressionParser` previously recursed once for every `Not` or prefix sign, and recursively parsed parentheses without a recovery limit. Long directives could throw `RangeError` during expression evaluation, constant indexing and activity tracking. This repair iterates flat unary chains and applies signs from the innermost operator outward using the existing token array. Parenthesis nesting uses the shared `MAX_EXPRESSION_DEPTH = 256`; excessive nesting leaves the expression unknown.

The existing coercions, comparison precedence, left-associative exponentiation and one optional exponent sign remain unchanged. No persistent cache is added. Expressions deeper than 256 levels now deliberately remain unknown even when a particular runtime could previously evaluate them.

Validation against `01f3f0db9b729df8e0c9191073a7174855bc6f2e`:

- Type checking passes.
- 34 new independent regressions and compatibility controls; baseline fails 11, repaired code passes all 34.
- 96 focused tests pass (2 skipped); 15,034 full-suite tests pass across 769 files (33 tests / 7 files skipped), with two workers, 276.19 seconds.
- 3,300 synthetic expression comparisons cover unary operators, nested parentheses below the limit, comparisons, exponent signs, Boolean and numeric coercions, Null/Empty/Nothing, dates and unknown names: no differences.
- 8,256 oracle sources across LF, CRLF and CR produce 24,768 complete conditional index and member-activity comparisons: no differences. This corpus does not substitute for the new independent deep-expression regressions.

Reproduce timing with:

```powershell
node scripts/benchmark-conditional-expression-depth.mjs --baseline=01f3f0db9b729df8e0c9191073a7174855bc6f2e
node scripts/benchmark-conditional-expression-depth.mjs
```

Node v24.18.0, AMD Ryzen 7 9800X3D. Two baseline and two repaired runs in ABBA order; each uses three warmups and nine measured calls. Source creation and bundling are outside the clock; complete expression evaluation including tokenization is inside. Outputs or the baseline's expected RangeError are independently checked on every call.

| Input | Baseline median ms | Repaired median ms |
| --- | ---: | ---: |
| 1 Not | 0.0093–0.0094 | 0.0093–0.0098 |
| 100 Not | 0.0556–0.0613 | 0.0584–0.0589 |
| 1 sign | 0.0026–0.0027 | 0.0025 |
| 100 signs | 0.0503–0.0577 | 0.0418–0.0455 |
| 1 parenthesis | 0.0022 | 0.0023–0.0024 |
| 100 parentheses | 0.0583–0.0601 | 0.0606–0.0691 |

All 12 calls in each baseline run throw for each 20,000-depth family. Repaired medians are 3.4815–3.5535 ms for Not, 1.8988–1.9463 ms for signs and 2.6828–2.7924 ms for parentheses (unknown recovery). Timings of failed baseline calls do not establish a throughput improvement. Ordinary-input costs are mixed, including slightly slower parenthesis controls. This is a bounded stack-safety repair, with no full-analyzer, heap-byte or editor-latency claim.
