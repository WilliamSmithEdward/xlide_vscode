# Interface declared-name spans

Implement Interface renamed signatures with a dynamically constructed regex ending in JavaScript's ASCII word boundary. Unicode names ending in non-ASCII letters and bracketed procedure names were left unprefixed, so the new procedure did not implement the named interface member. Bracketed public fields also lost their written escaping.

Use the parser's declared-name span, rebased to the copied signature after trimming/access-modifier removal. Synthetic field property headers carry their known name span and preserve bracketed spelling. Replace only that span with the interface-prefixed name. Parameters, defaults, comments, type suffixes and continued header text remain intact. Incomplete or out-of-header name spans retain the existing copied header rather than guessing a replacement. The refactor's regex-escaping dependency and dynamic name regex construction are removed.

Seventy new tests cover Greek, CJK, Thai, Devanagari with combining marks, bracketed keyword/space names, Sub/Function/Property Get, scalar/object field legs, LF/CRLF/CR, matching parameter/literal/comment text, type suffixes and default access with leading whitespace. Sixty-nine cases fail on baseline: 67 exact-result cases plus two deterministic constructor-work cases. Generating 1,000 ASCII members previously built 1,000 declaration-name regexes; the repair builds zero. All 70 pass again after restoring the exact tested source. Types, 104 focused tests and the full suite pass: 14,829 tests, 33 skipped, 761 passing files and seven skipped files, 72.32 seconds.

All 24,768 complete corpus queries across 8,256 sources and three newline styles match baseline. The corpus does not contain the reproduced failing name shapes; the independent new test expectations establish those corrections. This does not claim that arbitrary malformed interfaces compile in Office.

## Reproduction and scope

Run `node scripts/benchmark-interface-declared-names.mjs` in baseline/current/current/baseline order, supplying `--baseline=9e02c845cdebe03b949d32e0ef7c826ffaeb0428` on baseline runs. Node 24.18.0, AMD Ryzen 7 9800X3D; three warmups and nine measured rounds, ten complete refactor queries per round. Both ASTs/source construction are outside the clock; complete results are independently checked afterward.

ASCII procedure and unbracketed field controls have identical output. Unicode/bracketed baseline outputs deliberately check the known wrong names, so those timings are not equal-output comparisons. No cold parser, heap-byte, editor command or renderer-paint measurement is claimed.

Milliseconds per complete query, 1,000 declared members. Ranges are medians from two independent runs; fields generate 2,000 property stubs.

| Interface members | Baseline | Repair |
| --- | --- | --- |
| ascii | 1.12554–1.18164 | 0.31754–0.33121 |
| unicode | 1.23467–1.28548 | 0.47200–0.48492 |
| bracketed | 1.18433–1.21755 | 0.38051–0.39933 |
| fields | 2.21031–2.25267 | 0.54235–0.55575 |
