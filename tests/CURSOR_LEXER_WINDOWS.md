# Fresh-source cursor lexer work

`completionLineCursorContext` previously projected a short token window from
`tokenizeCached(source)`. The projection was bounded, but a changed source still
required the full-module lexer cache to compare or tokenize that module. Type
completion made the same request before member and identifier casing queries.
Bounding only the type query moved the full-stream cost to the next resolver.

Line queries now lex the current logical group and the preceding logical
boundary. They preserve the leading newline, token spans, trivia, contextual
keywords, and truncated-token behavior. Underscore-ended physical rows are
included conservatively so code and comment continuations retain their lexical
state. CR, LF, CRLF, a cursor inside CRLF, and lexer whitespace are covered.
Absolute line numbers are calculated once per window when a consumer requests
them; typing resolvers use spans and grammar instead. The full-prefix token API
retains its original metadata and cache contract.

Type detection lexes successive logical chunks until it has five grammar tokens,
preserving its existing behavior across newlines and comments. After 32 chunks,
a sparse tail uses the shared full-stream API. This avoids independently
re-lexing thousands of blank or comment rows. Long logical lines and sparse
lookbacks can still require substantial source work; this change does not claim
constant work for every malformed source.

## Regression evidence

`cursorLexerWindowWork.test.ts` compares complete token metadata and cursor
decisions with the unchanged full-prefix API at every offset of public lexical
fixtures under all three line endings. It checks fresh source revisions near
both ends of a large synthetic module, positive declaration and qualified type
positions, comments, continuations, sparse-tail fallback, absolute line numbers,
and JSON serialization. The cold-query tests count actual lexer input and
full-stream misses rather than warming that cache first.

Additional local checks compared 47,272 complete cursor contexts and 6,770 type
answers with frozen previous implementations. All agreed. The cursor checks
include generated malformed source with mixed line endings and unusual lexer
whitespace. A further 652 sampled private-module contexts agreed, including all
token and trivia fields. Private source and detailed logs remain ignored.

A 55,401-character cold type fixture previously missed the full-stream cache
with all 55,401 characters; the changed resolver lexed a 21-character chunk.
A separate 52,821-character cold line-context fixture lexed a 29-character
window. The public regressions use fresh revisions and reject full-stream
misses on ordinary and declaration queries near both ends.

## Controlled complete casing comparison

A local benchmark bundles frozen previous source and changed source with the
same dependencies. It asks the complete casing resolver to correct three
identifier occurrences in a public probe appended to a 969,328-character
private class. Each trial uses 15 warmups, 30 warm requests and 30 source
revisions that change a fixed-width comment. Four trials alternate variant
order. Every fresh request returns all three expected edits.

Fresh-request medians were 6.64–6.84 ms before and 4.91–5.32 ms after. Across
those trials, fresh p95 values ranged from 7.00–10.46 ms before and 5.54–10.51 ms
after; maxima ranged from 9.54–12.17 ms before and 7.19–10.80 ms after. Warm
medians were 0.59–0.69 ms before and 0.69–0.80 ms after. Window lexing adds small
warm-request allocations while removing the shared full-stream dependency from
ordinary fresh requests.

An earlier type-only prototype barely changed the complete casing median
(6.96 ms before, 6.90 ms after): the next resolver still requested full-module
tokens. The shared line-context change is necessary to remove that work from
the whole request. These are component measurements, not popup or hover
painting measurements.

## Editor validation

Compilation passed. The full unit suite passed 15,108 tests across 773 files,
with 33 tests and seven files skipped. The focused compatibility suite passed
390 tests across 20 files, with three opt-in benchmark tests skipped.

Pinned VS Code 1.139.1 passed 29 completion/Enter surface checks and two type
lookback checks. The private large-class harness passed all 12 tests, including
64 fresh-source renderer cycles and four real mouse hovers. Visible latency
median/p95/maximum was 16/32/32 ms for Backspace, 124/138/155 ms for the restored
menu, 31/32/36 ms for typing, and 1/1/11 ms for clearing the unmatched menu.
Mouse hover samples ranged from 337 to 381 ms. These include renderer painting,
CDP polling, and VS Code's normal completion/hover delays. They are separate
from provider response timing and do not isolate the UI effect of this change.
