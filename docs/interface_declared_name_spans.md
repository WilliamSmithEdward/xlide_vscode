# Interface declared-name spans

Implement Interface renamed signatures with a dynamically constructed regex ending in JavaScript's ASCII word boundary. Unicode names ending in non-ASCII letters and bracketed procedure names were left unprefixed, so the new procedure did not implement the named interface member. Bracketed public fields also lost their written escaping.

Use the parser's declared-name span, rebased to the copied signature after trimming/access-modifier removal. Synthetic field property headers carry their known name span and preserve bracketed spelling. Replace only that span with the interface-prefixed name. Parameters, defaults, comments, type suffixes and continued header text remain intact. Incomplete or out-of-header name spans retain the existing copied header rather than guessing a replacement. The refactor's regex-escaping dependency and dynamic name regex construction are removed.

Seventy new tests cover Greek, CJK, Thai, Devanagari with combining marks, bracketed keyword/space names, Sub/Function/Property Get, scalar/object field legs, LF/CRLF/CR, matching parameter/literal/comment text, type suffixes and default access with leading whitespace. Sixty-nine cases fail on baseline: 67 exact-result cases plus two deterministic constructor-work cases. Generating 1,000 ASCII members previously built 1,000 declaration-name regexes; the repair builds zero. All 70 pass again after restoring the exact tested source. Types, 141 focused tests and the full suite pass: 14,882 tests, 33 skipped, 763 passing files and seven skipped files, 72.78 seconds.

All 24,768 complete corpus queries across 8,256 sources and three newline styles match baseline. The corpus does not contain the reproduced failing name shapes; the independent new test expectations establish those corrections. This does not claim that arbitrary malformed interfaces compile in Office.

## Reproduction and scope

Run `node scripts/benchmark-interface-declared-names.mjs` in baseline/current/current/baseline order, supplying `--baseline=878faab4b92f9e45b0f022cfb59dff04ea5a50fa` on baseline runs. Node 24.18.0, AMD Ryzen 7 9800X3D; three warmups and nine measured rounds, ten complete refactor queries per round. Both ASTs/source construction are outside the clock; complete results are independently checked afterward.

ASCII procedure and unbracketed field controls have identical output. Unicode/bracketed baseline outputs deliberately check the known wrong names, so those timings are not equal-output comparisons. No cold parser, heap-byte, editor command or renderer-paint measurement is claimed.

Milliseconds per complete query, 1,000 declared members. Ranges are medians from two independent runs; fields generate 2,000 property stubs.

| Interface members | Baseline | Repair |
| --- | --- | --- |
| ascii | 1.17507–1.21297 | 0.38285–0.41188 |
| unicode | 1.36622–1.40459 | 0.58777–0.61972 |
| bracketed | 1.31573–1.35626 | 0.53241–0.54168 |
| fields | 2.22035–2.22634 | 0.60051–0.67371 |

Single-member controls remain in the microsecond range: ASCII 0.00361–0.00395 → 0.00280–0.00306 ms, Unicode 0.00320–0.00339 → 0.00289–0.00302 ms, bracketed 0.00316–0.00321 → 0.00306–0.00318 ms, fields 0.00373–0.00375 → 0.00265–0.00272 ms.
