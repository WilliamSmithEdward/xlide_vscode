# Unreachable-code removal boundaries

The Remove unreachable code action previously widened a dead statement to its whole physical line. For example, removing Debug.Print from Exit Do: Debug.Print "dead" also removed the live Exit Do, leaving an unconditional Do/Loop. Issue #945 records the reproduction.

The edit now removes the dead statement range and its separating colon while retaining live text at both boundaries. Whole standalone dead lines still include their line ending. Comments within the removed range are retained using lexer comment tokens, so quote characters inside strings are not treated as comments. LF, CRLF and CR are preserved. The declaration guard checks the actual edit range: a retained same-line Dim no longer blocks a safe assignment removal, while a dead block containing a declaration still has no removal action.

This changes removal metadata, not unreachable-code detection. It does not execute VBA or claim a performance improvement.

## Validation

- Eighteen new public diagnostic/code-action tests cover shared-line Exit Do, multi-line dead ranges, trailing and interior comments, landing labels, standalone lines and declaration guards across three line endings. The baseline fails thirteen and passes five controls.
- Fifty-eight focused tests and the TypeScript check pass.
- The full suite passes: 667 files, 13,417 tests passed and 15 skipped, on base d3c0f93b.
- Comparison against that baseline covers 8,256 corpus sources and 75 generated sources. Complete diagnostic arrays match in 8,251 cases. Eighty cases intentionally change only unreachable-removal metadata; diagnostic detection and all other metadata match in every case. Cached lexer tokens and their trivia are frozen during comparison, with no uncaught exceptions.

Of the five corpus payload changes, one enables the safe assignment removal beside a retained Dim. Four affect removal boundaries in reserved-word-label fixtures already rejected by the VBA oracle. The seventy-five generated payload changes exercise terminators, layout and line endings.
