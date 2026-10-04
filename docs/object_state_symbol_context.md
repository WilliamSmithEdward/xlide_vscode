# Object-state walks belong to bound symbol snapshots

`objectLetStateAt` asks a shared procedure walk whether a default-member Let
assignment reaches an object that is set, unset or unknown. The walk reads local,
implicit and module-variable facts from bound symbols. Its old cache was weakly
keyed only by the parsed procedure, with source, conditional activity and member
completion context checks. It omitted the bound-symbol snapshot.

A retained parsed procedure could therefore reuse another snapshot's result.
For `Dim value As Object: value = 3`, a normal local is unset. A bound snapshot
marking the local auto-instantiated or Static excludes it from that tracking,
so the answer is unknown. Fresh member-context controls establish those answers;
with a retained member context, the old cache incorrectly keeps returning unset.
Six regressions reproduce this across LF, CRLF and CR.

The cache now uses weak symbol ownership, then a weak procedure map. Source,
module identity, conditional activity and member completion context must match
within a symbol snapshot. This keeps symbol-dependent walks separate and lets
released snapshots be collected. Alternating retained snapshots keeps each warm:
a regression checks that 20 alternating queries reread no declaration facts.

The repair does not detect mutations inside an unchanged symbol snapshot or
member context; those identities represent stable analysis inputs. The tests
construct separate completed snapshots before querying them. No timing or
whole-editor latency improvement is claimed.

The linked issue and PR record final type, focused/full suite and frozen baseline
comparison counts. The baseline differential compares complete cached findings,
Let states and diagnostic arrays for unchanged contexts; the new regressions
assert the intentionally corrected behavior when bound contexts differ.

Type checking and 179 focused tests pass, including seven new object-state
context/reuse tests. The frozen baseline differential includes 8,256 corpus
sources and 36 generated cases: 13,482 diagnostics, no internal errors, and
identical complete object-walk findings, Let states and diagnostic arrays for
stable contexts.

The default full suite initially hit two unchanged timing limits: many-locals
ratio 12.25 against 9 and the inflate fixture test at 5 seconds. Both pass
unchanged in isolation (21 tests). The full unchanged suite then passes with
two workers: 737 files and 14,535 tests (7 files and 33 tests skipped).
