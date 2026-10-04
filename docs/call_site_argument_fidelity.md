# Call-site argument fidelity and bounded scanning

Issue: #960. Baseline: 301debbcb1354d6bcb8b4cbccb4f348c974648ad.

Introduce Parameter previously replaced stripped argument text. For example, appending 3 to Go "hello" produced Go  3"hello"; a colon-separated call could replace its neighboring statement; and Go (1) became Go (1, 3), changing a parenthesized bare argument into a call-list form.

The helper now returns insertion-only edits at the end of the actual argument list. Original strings, dates, comments, whitespace and nested calls remain intact. A query-owned logical-line token index handles statement separators, inline Else branches, named arguments, bracketed names and continuations. Parenthesis pairs and statement ends are indexed once; nested call insertions do not overlap. Introduce Parameter supplies the new formal parameter name when extending named arguments. Explicit zero-argument Call statements gain parentheses, as required by [Microsoft's Call syntax](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/call-statement). Assignment targets remain excluded.

Simple physical lines retain a fast path. Complex lines use uncached fragment lexing so their tokens cannot evict the shared module lexer cache. The index is owned by the current query and is discarded afterward; source tokens and trivia are not mutated. Tests guard at most one lexer pass per logical line with 1, 100 and 1,000 nested or colon-separated calls, bound total fragment input by source length, and verify that 1,000 unique fragments preserve the cached module token array.

## Validation

- The three new test files contain 118 tests: the baseline fails 86 and passes 32 controls. All 205 focused tests in six files pass with the fix, across LF, CRLF and CR where applicable.
- Type checking passes. The original validated base passes 676 files, 13,643 tests and 18 skips. After removing the squash-merged prerequisite and rebasing onto main a30644ff, type checking and the complete integration suite pass: 686 files, 13,735 tests and 22 skips. The tested call-site and Introduce Parameter source is unchanged by that rebase.
- 216 generated complete Introduce Parameter operations check primary and external module output against independently constructed expected source. All candidate outputs match; the baseline renders 84 correctly. These cover strings, nested calls, dates, named arguments, continuations, explicit Call, zero-argument expressions, bare parentheses and inline branches.
- An observational differential run covers 8,256 corpus sources and 19,814 complete helper queries: 17,654 complete metadata arrays match, 17,966 rendered sources match, and 1,848 rendered sources change. Changes include repaired syntax and changed edit representation; this is not a claim that every output is unchanged or that every corpus binding is proven correct. Candidate corpus edits have zero overlaps. Owned and cached lexer tokens and trivia are frozen during the run, with zero uncaught exceptions.
- No Office/VBA runtime or interactive tree latency measurement is claimed by these tests.

## Reproducible timing

Run node scripts/benchmark-call-site-arguments.mjs --baseline=301debbcb1354d6bcb8b4cbccb4f348c974648ad --rounds=9 for the baseline and omit --baseline for the candidate. Four sequential runs use baseline/candidate/candidate/baseline order, three warmups and nine measured rounds each. Environment: Node 24.18.0, Windows, AMD Ryzen 7 9800X3D. Timings below are the two median measurements in milliseconds. Only helper execution is timed; complete rendered output is checked separately against expected text. Fresh inputs vary a leading comment to miss the full-source stripped-text cache; this does not assert that every other cache is cold.

For 1,000 nested calls, correct output is identical and cached time falls from 15.945–16.2003 ms to 0.5448–0.5713 ms. Colon-separated output becomes correct and its quadratic repeated work disappears. Ordinary cached numeric calls become slightly slower (0.2017–0.2319 to 0.2776–0.2815 ms); bracketed calls also cost more at 1,000 calls (0.2085–0.2111 to 0.3869–0.4101 ms). Small and fresh cases are mixed. These measurements support the bounded nested-call improvement, not a universal speedup. Baseline output is correct in 26 of 36 configurations; candidate output is correct in all 36.

| calls/layout/cache | before ms | after ms | baseline correct | candidate correct |
| --- | --- | --- | --- | --- |
| 1/number/cached | 0.0073 / 0.0075 | 0.0064 / 0.0063 | yes | yes |
| 1/number/fresh | 0.0057 / 0.0082 | 0.006 / 0.0066 | yes | yes |
| 1/string/cached | 0.0018 / 0.0041 | 0.0018 / 0.0018 | no | yes |
| 1/string/fresh | 0.0051 / 0.0064 | 0.0048 / 0.005 | no | yes |
| 1/bracketed/cached | 0.0032 / 0.0035 | 0.0021 / 0.0021 | yes | yes |
| 1/bracketed/fresh | 0.0074 / 0.007 | 0.0044 / 0.0045 | yes | yes |
| 1/nested/cached | 0.0016 / 0.0035 | 0.002 / 0.0021 | yes | yes |
| 1/nested/fresh | 0.0038 / 0.0057 | 0.0051 / 0.0043 | yes | yes |
| 1/colon/cached | 0.0017 / 0.0022 | 0.0025 / 0.0025 | yes | yes |
| 1/colon/fresh | 0.0034 / 0.0057 | 0.0034 / 0.0035 | yes | yes |
| 1/comments/cached | 0.0022 / 0.0033 | 0.0018 / 0.0018 | yes | yes |
| 1/comments/fresh | 0.0042 / 0.0057 | 0.0041 / 0.0038 | yes | yes |
| 100/number/cached | 0.0493 / 0.0484 | 0.0715 / 0.0682 | yes | yes |
| 100/number/fresh | 0.0616 / 0.0561 | 0.0536 / 0.0839 | yes | yes |
| 100/string/cached | 0.0323 / 0.032 | 0.0243 / 0.0229 | no | yes |
| 100/string/fresh | 0.0488 / 0.0775 | 0.0399 / 0.0387 | no | yes |
| 100/bracketed/cached | 0.0361 / 0.0316 | 0.0533 / 0.0526 | yes | yes |
| 100/bracketed/fresh | 0.0494 / 0.0477 | 0.0663 / 0.0664 | yes | yes |
| 100/nested/cached | 0.2018 / 0.1995 | 0.1596 / 0.1564 | yes | yes |
| 100/nested/fresh | 0.1972 / 0.1988 | 0.061 / 0.0686 | yes | yes |
| 100/colon/cached | 0.5704 / 0.5706 | 0.1224 / 0.139 | no | yes |
| 100/colon/fresh | 0.589 / 0.6108 | 0.0687 / 0.0665 | no | yes |
| 100/comments/cached | 0.0309 / 0.0294 | 0.0283 / 0.0283 | yes | yes |
| 100/comments/fresh | 0.0493 / 0.048 | 0.0457 / 0.0451 | yes | yes |
| 1000/number/cached | 0.2319 / 0.2017 | 0.2776 / 0.2815 | yes | yes |
| 1000/number/fresh | 0.3712 / 0.4322 | 0.3355 / 0.33 | yes | yes |
| 1000/string/cached | 0.2066 / 0.2133 | 0.2927 / 0.2939 | no | yes |
| 1000/string/fresh | 0.3554 / 0.3503 | 0.3563 / 0.324 | no | yes |
| 1000/bracketed/cached | 0.2111 / 0.2085 | 0.4101 / 0.3869 | yes | yes |
| 1000/bracketed/fresh | 0.3686 / 0.3613 | 0.5806 / 0.5477 | yes | yes |
| 1000/nested/cached | 16.2003 / 15.945 | 0.5448 / 0.5713 | yes | yes |
| 1000/nested/fresh | 17.3695 / 16.282 | 0.5015 / 0.5321 | yes | yes |
| 1000/colon/cached | 55.9937 / 52.4132 | 0.5688 / 0.5576 | no | yes |
| 1000/colon/fresh | 53.1382 / 53.3553 | 0.5543 / 0.6574 | no | yes |
| 1000/comments/cached | 0.2471 / 0.229 | 0.2366 / 0.2077 | yes | yes |
| 1000/comments/fresh | 0.4204 / 0.4077 | 0.3947 / 0.3736 | yes | yes |
