# Formatter procedure stack position

Each block opener previously searched the open-block stack to decide whether it
was inside a procedure. At module level, a nested conditional stack has no
procedure, so every opener scans all earlier blocks. The 500-level work-count
case visits 124,750 stack entries before producing its formatted output.

The stack can contain only one procedure: a new procedure header removes the
previous procedure before pushing itself. Record that procedure's stack position
and use it for opener ownership. Stack truncation clears the position if it
removes the procedure; a replacement header records its new position. Opener
classification, indentation, spacing, keyword/identifier casing and token-stream
safety checks are unchanged. No source cache or public API was added.

## Reproduction

```powershell
npx vitest run tests/formatterProcedurePositionWork.test.ts tests/vbaFormatter.test.ts tests/vbaFormatterCorpus.test.ts tests/vbaFormatterTokenReuse.test.ts
node scripts/benchmark-formatter-procedure-position.mjs --baseline=9ccf1bc21f16057020be3123b28ea7fe9c79d123
node scripts/benchmark-formatter-procedure-position.mjs
```

Nine size/EOL work regressions fail on baseline and pass with no stack search.
They compare independently constructed complete expected indentation. Seven
output controls pass on both: module branch truncation, mismatched closers,
replacement procedure headers and stray End Sub boundaries, across three EOLs.

## Measurement

Node 24.18.0, Ryzen 7 9800X3D; nine samples after three warmups, separate process
order baseline/current/current/baseline. The actual whole formatter, including
both lexing passes and safety comparison, is timed. Source creation and complete
independent expected-output assertions are excluded. Tabs/spaces and inside/
outside-procedure cases cover depths 1/10/100/500.

| 500-level scenario | Baseline median ms | Fixed median ms |
| --- | ---: | ---: |
| Module-level, spaces | 3.5568–3.6757 | 3.1150–3.3738 |
| Module-level, tabs | 2.1578–2.2874 | 2.0565–2.1919 |
| Inside procedure, spaces | 2.5645–2.5691 | 2.5144–2.5172 |
| Inside procedure, tabs | 1.5484–1.7714 | 1.4666–1.4735 |

Shallow controls are mixed: at depth one with spaces, module-level medians were
0.0272–0.0381 before and 0.0315–0.0367 after. The spaced module-level output has
1,011,014 characters, and writing/lexing indentation still grows with depth.
This removes one quadratic search; it does not make the whole formatter linear
in source length, eliminate deep output costs or establish an editor-latency
improvement. Fixed maxima remain up to 4.9694 ms in the deep spaced module case.

## Validation

Types and 67 focused tests pass, including existing refusal/safety controls.
Frozen lexer-token/trivia baseline comparison covers 8,256 oracle sources plus
3,000 generated malformed block/procedure sequences across LF/CRLF/CR and four
indentation options. All 45,024 complete formatting results match. Full suite passed: 743 files, 14,614 tests, 33 skipped tests, 77.92 seconds.
