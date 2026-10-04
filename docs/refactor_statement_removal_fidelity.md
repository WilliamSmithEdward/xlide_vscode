# Refactoring statement removal fidelity

Inline Variable and Introduce Parameter widened both removed statements to complete LF-based physical lines. Valid colon-separated neighbors were erased; CR-only input could yield overlapping whole-module removals and an empty result. Extract Method already carried a more precise removal routine and deletion coalescer. Those routines now live in refactor/shared.ts and serve all three refactors. The shared removal keeps neighboring statements and labels, preserves trailing comments through lexical code-span extraction, and merges overlapping/adjacent deletions without mutating input spans. Assigned-value extraction excludes trailing lexical comments rather than treating them as part of the initializer.

wholeLineSpan now delegates to the existing CR-aware physical-line implementation while preserving its former LF/CRLF span-start behavior when the offset is inside a CRLF pair. The established nearest-break helper itself is unchanged. Existing callers in Encapsulate Field and dead-code diagnostics were inspected; physical-line widening alone does not repair caller-specific statement/documentation mistakes. Actual diagnostic probes uncovered remaining issues #945 (unreachable removal erases a live Exit Do) and #946 (module declaration removal uses LF-only attached-comment bounds on CR source). Those issues remain separate work, not claimed fixed here.

Validation on baseline 2191273f:

- 51 new exact-text/edit-boundary tests: 49 fail before the fix and two LF/CRLF boundary controls pass. Both refactors cover following/preceding neighbors, labels, trailing comments, quoted colons, declaration/assignment/use on one physical line and adjacent deletion spans across LF, CRLF and CR.
- 145 focused tests include existing extraction/grouped-declaration behavior and nearest-break contracts, including fractional and non-finite offsets.
- Type checking passed. The final full suite passed 664 files / 13,419 tests (14 skipped).
- 1,800 complete compatible Inline Variable/Introduce Parameter results match baseline on LF/CRLF scalar/object, atomic/compound value, refusal and local call-site fixtures. Rendered source is compared exactly. Edit ordering and equivalent adjacent/overlapping deletions are canonicalized before metadata comparison because the shared coalescer intentionally changes deletion representation. No exception results occurred. Intentional neighbor/CR/comment repairs have independent exact-text expectations in the regression tests.

These are source-text refactoring and diagnostic-edit observations, not Office/VBA execution. No performance speedup is claimed; this removes duplicate deletion logic and fixes destructive edits. The entire analyzer/repository audit remains incomplete.
