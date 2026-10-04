# Conditional inactive-arm evaluation gates

`applyConditionalDirective` previously evaluated every `#If` and `#ElseIf` expression, even under a proven-inactive parent or after an earlier arm had already been taken. The resulting activity could no longer change, so tokenization, expression parsing and compiler-constant setup were wasted.

The repair treats nested conditions under an inactive parent as inactive and returns early for `#ElseIf` when its parent is inactive or the chain has already seen a true arm. Every `#Const` still replays, including definitions in skipped arms. Uncertain parents and later arms following false conditions still evaluate. The same gates serve tracker construction, offset replay and Null-condition scanning. No persistent cache is added.

Validation against `60b20fa0bfe5326a12c3ef0b26beb0fd1b948da0`:

- Type checking passes; 72 focused tests pass, 2 skipped.
- Ten new work/semantic controls: baseline fails five work assertions, repaired code passes all ten. After module parsing, 1,000 skipped ElseIf or nested inactive conditions make 1,001 lexer calls before and one after. Constant replay, unknown activity, fresh environments and definitely evaluated Null conditions have independent assertions.
- Full suite: 770 files / 15,029 tests pass, 7 files / 33 tests skipped, 72.68 seconds.
- 8,256 oracle sources across three physical line-ending forms and two compiler environments yield 49,536 complete comparisons with no differences. Each compares constant indexes, Null conditions, tracker activity, offset activity replay, full `analyzeModule` diagnostics and internal failures. This does not replace the independent skipped-arm work controls.

Reproduce:

```powershell
node scripts/benchmark-conditional-inactive-gates.mjs --baseline=60b20fa0bfe5326a12c3ef0b26beb0fd1b948da0
node scripts/benchmark-conditional-inactive-gates.mjs
```

Node v24.18.0, AMD Ryzen 7 9800X3D. ABBA order, two baseline and two repaired runs, three warmups and nine measured rounds. Each arm expression has 100 Not operators. Source creation and AST warming are outside the clock; complete tracker construction or complete `analyzeModule` diagnostics are inside. Every call independently verifies activity or the complete empty diagnostic output plus the absence of internal failures. The uncertain control requires all arm expressions to evaluate.

| Scenario | Baseline median ms | Repaired median ms |
| --- | ---: | ---: |
| Tracker, 1 settled ElseIf | 0.0658–0.0661 | 0.0121–0.0124 |
| Tracker, 1 inactive nested If | 0.0324–0.0462 | 0.0045–0.0046 |
| Tracker, 1 uncertain arm | 0.0308–0.0314 | 0.0311–0.0314 |
| Tracker, 1,000 settled ElseIf arms | 11.6199–11.7422 | 0.0486–0.0537 |
| Tracker, 1,000 inactive nested Ifs | 11.2281–11.5756 | 0.0530–0.0553 |
| Tracker, 1,000 uncertain arms | 11.3028–11.4159 | 12.5430–12.7268 |
| Full diagnostics, 1 settled ElseIf | 0.6277–0.6382 | 0.5442–0.5443 |
| Full diagnostics, 1 inactive nested If | 0.4037–0.4095 | 0.3846–0.4115 |
| Full diagnostics, 1 uncertain arm | 0.4703–0.4748 | 0.4461–0.4643 |
| Full diagnostics, 1,000 settled ElseIf arms | 49.8636–49.8773 | 15.1129–15.5482 |
| Full diagnostics, 1,000 inactive nested Ifs | 54.5776–55.6958 | 20.4340–20.6407 |
| Full diagnostics, 1,000 uncertain arms | 50.5310–50.9140 | 53.5665–53.9195 |

The measured skipped-arm path improves. The 1,000-arm uncertain control is slower (about 1.1–1.4 ms for tracker construction and 2.7–3.4 ms for full diagnostics); those expressions cannot be skipped. Small-input costs are mixed. This is not a universal analyzer speedup and does not establish cold-parser, retained-heap, worker or renderer latency improvements. Compiler lookup setup still occurs per evaluated expression and remains a separate audit candidate.
