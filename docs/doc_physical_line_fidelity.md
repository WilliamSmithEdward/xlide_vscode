# Documentation physical-line boundaries

Issue #946 reproduces an unused module declaration removal that also deletes preceding Option Explicit from CR-only source. The whole-line deletion boundary was correct, but attachedCommentsStart replaced its start with an LF-only scan returning zero.

The documentation scanner now treats CR, CRLF and LF as physical terminators. leadingDocLines and attachedCommentsStart share a bounded backward scan, keeping XML doc lines and XLIDE directives attached in either order and stopping at ordinary comments or blank lines. Header scanning uses the same physical line endings. Text and tag spans retain source coordinates.

The caller audit also repaired Move to Module's own LF-only removal and string/comment bounds. It retains preceding code and following procedures, moves only attached documentation, and preserves CR-only serialization. Existing LF/CRLF serialization policy remains in place. Encapsulate Field now sees CR documentation through the corrected shared scanner, so the description moves once to the generated property.

Recognizing CR documentation exposed LF-only code-action bounds too. Adding a param tag previously inserted it at module EOF; whole-tag removal could consume following code. The documentation actions now use physical line bounds for insertion, deletion and a preceding next-line directive, and use CR for generated lines in CR-only source.

## Validation

Forty new tests exercise actual analyzer diagnostics and public code actions, documentation extraction and spans, module headers, blank/comment boundaries, procedure moves, qualified calls in code versus strings/comments, and field encapsulation across LF, CRLF, CR and mixed endings. Baseline 320ff50b fails nineteen and passes twenty-one controls. The focused five-file suite passes 135 tests, and the type check passes.

Complete diagnostic, symbol and documentation results match the baseline for all 8,256 corpus sources and 90 compatible generated sources. Forty-five CR-only generated sources match canonical LF results after mapping source offsets and line endings. Forty-five complete Move to Module results match the baseline or canonical LF as appropriate. Parsed ASTs and cached lexer tokens/trivia are frozen; no uncaught exceptions occur.

This is a source-fidelity correction and removes duplicated backward-line traversal. No runtime speedup is claimed. The full suite passes on base 320ff50b: 671 files, 13,520 tests passed and 17 skipped.
