# Continued initializers in local refactors

A valid assignment such as `limit = _` followed by `3` on the next physical line was parsed as one statement, but the shared initializer reader used a regular expression whose value capture could not cross that line break. Both Inline Variable and Introduce Parameter refused the operation with “Could not read the value assigned to 'limit'.”

The shared reader now folds lexer-recognized continuation trivia into whitespace before applying its existing assignment match. It does this only for multiline statements, preserving string/date token contents and ordinary logical newlines. It retains the existing assignment/comment boundaries; public refactors remove the complete original continued statement using its unchanged span. Single-line matching remains on the original path.

Validation: 29 new regression/control tests cover continuation positions, Set, complete inline/parameter edits, string/comment text, uncontinued newlines and ordinary matching across LF, CRLF and CR. Eighteen fail against baseline; eleven already pass. Type checking, 86 focused tests and the final full suite pass: 776 files / 15,193 tests, with 7 files / 33 tests skipped, in 73.66 seconds.

A frozen comparison against `b8630fb7dde62e1bfa645a3b1a4e203f3f926013` runs both public refactors with single-line target initializers in all 8,256 oracle source environments across three line endings. All 49,536 queries succeed and preserve exact result objects, edited source and cross-module edit lists. Intended continued behavior is independently asserted by the new tests; the comparison is compatibility evidence for the unchanged target path, not a claim that every oracle refactor is semantically valid.

This is a correctness repair discovered during the performance audit. It makes no timing, heap or editor responsiveness claim. The separately suspected growing-array initializer-name search remains to be measured once valid continued initializers reach that path.
