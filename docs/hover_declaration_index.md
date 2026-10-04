# Indexed local hover declarations

Bare-symbol hover already borrows an immutable, bounded editor symbol snapshot.
It nevertheless rescanned all flattened declarations for the enclosing procedure,
then procedure children, module declarations and enum members on each request.
Ten requests with 1,000 preceding procedures read 10,030 original entries;
with 1,000 preceding locals they read 20,030. Six work-count regressions reproduce
the repeated scans at 10, 100 and 1,000 declarations.

A weakly owned index now records procedure ranges, first module declarations and
first enum members. Each encountered procedure lazily receives its first-name
local map. Normal ordered procedure ranges use binary lookup by inclusive end;
overlapping or unordered ranges retain the original first-match search over
procedures. Local/module/enum precedence and first matching declaration wins are
unchanged. Results and external documentation are still constructed on every
request. The index follows the lifetime of the borrowed immutable symbol graph.

## Measurement

Run `node scripts/benchmark-hover-declarations.mjs`, or add
`--baseline=0e0882d172119f13f9bbc434394c5bac98396bd1` for the old implementation.
The script independently checks complete expected hover records. Three warmups
precede nine rounds. Warm rows measure 100 complete repeated hover requests after
parsing, symbol binding and first indexing; cold-index rows measure one hover in
the first procedure after symbol binding, including index construction.
Source construction is excluded. Paired before/after/after/before runs on
Node 24.18.0 / Ryzen 7 9800X3D gave:

| 1,000 declarations | Before median ms | After median ms |
| --- | --- | --- |
| 100 warm hovers, preceding procedures | 0.6805–0.6954 | 0.1950–0.1956 |
| 100 warm hovers, preceding locals | 1.5276–1.5302 | 0.1896–0.1920 |
| One first-scope hover, cold index | 0.0505–0.0536 | 0.1389–0.1395 |

The first index adds about 0.085–0.089 ms in this cold control. Small modules
and one-off requests need not improve. These are analyzer component measurements,
not editor or transport latency. Malformed overlapping ranges retain linear
procedure lookup; the expensive original local/root/enum rescans remain indexed.

## Validation

Type checking and 68 focused tests pass (two existing opt-in benchmarks skipped).
Eleven new cases cover work bounds, first-name and local/module/enum precedence,
LF/CRLF/CR, inclusive touching ranges and overlap fallback. Existing tests cover
source/module-kind changes, fresh output and external documentation updates.
The frozen baseline comparison spans 8,256 corpus sources plus 36 generated
cases: 234,754 positions, 98,968 resolved hovers, no exceptions, and identical
complete HoverInfo/undefined results. AST, lexer tokens and trivia were frozen.

The full suite passes: 737 files and 14,539 tests (7 files and 33 tests skipped).
