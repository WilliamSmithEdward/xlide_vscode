# Sparse reference statement classification

`classifyReferenceKinds` previously traversed every cached token and built each
statement segment even when a query requested two occurrences in one short
statement. This is used by references, diagnostic dead-code checks and the
local refactor helpers. Token caching avoided lexing again, but the unrelated
token sweep and segment allocations remained.

For sparse queries, sort unique finite requested offsets into source order,
binary-search exact token starts, and classify only complete statements that
contain a requested token. A previously covered statement is skipped for later
queries in it. Statement boundaries remain lexer newline/colon tokens, preserving
line continuations, labels, comments, and malformed-input behavior. Unmatched
positions still receive the read default in original request order.

Dense queries (requested-offset count at least one eighth of token count) keep
the original streaming pass. A sparse query that needs 64 backward steps to find
its statement context also falls back to streaming: this avoids traversing a
long statement backward and forward. Reset earlier sparse results before that
fallback to preserve complete Map insertion order. The streaming implementation
is shared by both fallbacks; there is no additional source or token cache.

## Validation

Baseline `6059d89dc95f54d986a16eadea77d360859b2a20`: twelve failing work
regressions and two passing result controls. Fifteen new tests cover beginning,
middle, end and separated statements with LF/CRLF/CR, unrelated token-kind work,
Map ordering, duplicates, nonfinite/unmatched offsets, continued inline If and
colon neighbors, and fallback after previously classified statements. The
baseline reads roughly 10,000 unrelated token kinds in the 1,000-statement
fixture; fixed queries read at most six boundary token kinds.

Type checking and 87 focused tests passed. Full suite: 745 files passed,
14,641 tests passed; 7 files / 33 tests skipped (85.47 seconds).

A differential across 8,256 oracle sources plus 1,500 generated malformed and
mixed-statement sources compared 87,804 complete ordered Maps. Cached token
arrays, tokens and trivia were frozen; all matched. Twelve additional complete
Map comparisons cover long statements at 64/65/1,000/10,000 arguments, earlier
classified statements, LF/CRLF/CR and default offsets; all matched with frozen
tokens.

## Measurement

Run `node scripts/benchmark-reference-sparse-statements.mjs`, optionally with
`--baseline=6059d89dc95f54d986a16eadea77d360859b2a20`.
Separate baseline/fixed/fixed/baseline processes on Node 24.18.0 / Ryzen 7 9800X3D
used three warmups, nine measured rounds and batches of twenty, reporting
per-call averages. Complete classification and token-cache retrieval are timed;
lexing is primed in warmups, fixture construction and independent complete Map
assertions are excluded.

| Shape (10,000 statements/arguments) | Baseline median ms | Fixed median ms |
| --- | --- | --- |
| First statement, two offsets | 0.67736–0.67868 | 0.000625–0.001085 |
| Middle statement, two offsets | 0.55319–0.56237 | 0.000635–0.001160 |
| Last statement, two offsets | 0.69793–0.70216 | 0.000640–0.001255 |
| First + last, four offsets | 0.72916–0.73602 | 0.001025–0.001605 |
| Dense, 20,000 offsets | 2.41348–2.44903 | 2.38127–2.47702 |
| One long statement, last argument | 0.20770–0.24869 | 0.20255–0.21895 |

Dense and long-statement controls have small mixed changes. The one-statement
first-query control is 0.000650–0.000655 to 0.000680–0.000705 ms; other small
cases are mixed. Fixed maximum batch averages were 0.00249 ms for the separated
case, 2.66065 ms for dense and 0.37125 ms for the long statement.

Cold lexing, dense queries and the long-statement fallback still visit the whole
module. These are warm classifier measurements, not complete refactor or editor
latency results.
