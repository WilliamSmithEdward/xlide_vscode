# Reusing immutable module snapshot facts

## Profile finding and change

The 800-cycle large-class capture after the provider cursor-window change spent
6.22 seconds in DefType source scans and 3.35 seconds visiting procedure bodies
for conditional-directive presence. Multiple symbol and semantic projections
ask the same questions about unchanged source.

DefType facts now have an eight-source value cache. A hit adopts the caller's
source instance. Every projection receives its own map, preserving ownership
if a consumer casts the public ReadonlyMap and mutates it. Source changes and
cache eviction still rescan with the existing regular expression.

The parser already knows when no directive tokens occur. It now records that
fact in a WeakSet of its immutable module snapshots, allowing conditional
queries to return without revisiting every body. Trees with possible directives
and caller-supplied trees retain the existing AST traversal. The weak metadata
retains no old trees. The uncached reference parser supplies an independent
traversal for differential checks.

## Regression evidence

Two work-count cases previously scanned the same source three times across
three differently named symbol projections; they now scan it once. Two queries
previously read 2,400 procedure bodies in a 1,200-procedure module; they now read
zero. Tests also cover map ownership, changed DefType text, eviction, directives
in bodies/Enum/Type, malformed input, comments/strings/dates, adding/removing a
directive in a large module, and caller-supplied mutable trees.

Compilation and 68 targeted tests passed. Before rebasing onto main's constant
dependency change (#1105), the full suite passed 14,693 tests across 752 files,
with 33 skipped tests and four workers. Compilation and the targeted suite
passed after the rebase too.

## Controlled private component comparison

The private class is known directive-free. On a 969,408-character fixture,
comparison bundles share all dependencies and differ only in the three source
files implementing these snapshot changes. One hundred warm directive queries
measured median/p95/max 0.3545/0.3682/0.4059 ms before and
0.0001/0.0002/0.0328 ms after. The work-count regression is the stronger evidence
for the very small cached-query timing.

Thirty warm batches of three symbol projections measured 6.26/7.43/7.52 ms
before and 1.75/2.55/2.74 ms after. Thirty fresh-source parse-and-three-projection
batches measured 8.29/10.67/12.63 ms before and 5.26/8.31/10.15 ms after. Each
fresh batch changed a fixed-width synthetic comment. These measurements do not
establish menu painting latency or eliminate the remaining popup outliers.
Private source and bundles stay in ignored .vscode-test.

## Renderer validation

The attempted 1,000-cycle capture completed 813 member-recovery cycles and
50 mouse hovers, then failed while waiting for the fresh synthetic nonce after
cycle 812. It did not fail during Backspace or menu recovery. Actual document
metadata showed the caret on the preceding synthetic statement at character
25, while the member line still ended in .cez. The failure lacked the actual
nonce and preceding-row/focus diagnostics needed to distinguish a missed test
predicate from an editing/rendering problem. Both profiles were retained.
This capture is incomplete and is not counted as a passing long-run validation.
Further diagnosis and sustained validation are required before making the draft PR ready for merge.

The unprofiled repeat failed at Backspace cycle 729. Failure diagnostics showed
the caret on the following line (26445, character 6), while the expected member
line was now 25 characters instead of the original expression. The preceding
nonce was visible and focus remained on the native edit context. This is
consistent with a preceding navigation/edit step operating on the wrong line;
it does not establish a blocked Backspace command. This repeat is also not
counted as a pass. Navigation diagnostics and partial sample retention are being
added before changing the probe's behavior or attributing the failure.

The opt-in XLIDE_PERF_NAV_DIAGNOSTICS capture records only caret-row matching,
row lengths, focus and widget counts after each synthetic navigation key.
Partial samples are retained even when the capture fails. No private source or
module labels are recorded.

A repeat with navigation diagnostics and both CPU profiles enabled passed all
12 integration cases, including 1,000 fresh-source recovery cycles and 62 mouse
hovers. Median/p95/max observations were Backspace 16/35/53 ms, menu recovery
117/139/187 ms, typing 16/45/75 ms, miss clearing 1/13/51 ms and hover
375/406/417 ms. The additional observations alter timing; this is diagnostic
validation, not an isolated before/after UI comparison. Profiles and the exact
bundle/source map were retained locally.

The retained navigation sequence shows that cursor DOM geometry can still
reflect the old row after navigation events have been dispatched. Fresh-source
setup now waits for the visible caret on the preceding statement, its End and
Shift+Home positions, then the return to the member expression and its end.
The nonce check verifies the entire synthetic statement on the row immediately
above the member expression, ignoring legitimate casing and display whitespace.
These are setup conditions outside the measured Backspace/typing response;
there are no new product waits. This prevents a wrong-line synthetic edit from
being reported later as a missing Backspace paint. Compilation and the 64-cycle
short run passed all 12 cases with these stronger checks. The sustained unprofiled
repeat of the final probe passed too.

The final unprofiled capture passed all 12 integration cases over 1,000 fresh
sources and 62 hovers. Median/p95/max observations were Backspace 18/34/67 ms,
menu recovery 118/140/206 ms, typing 16/47/61 ms, miss clearing 1/12/42 ms and
hover 379/409/418 ms. Its 48 command-based typing/deletion pairs measured median
1.85/1.90 ms and maximum 9.32/13.87 ms; the separate host heartbeat maximum was
27.39 ms. The slowest menu samples were cycles 197 (206 ms), 216 (198 ms) and
190 (195 ms). No wrong-row edit or recovery failure occurred. The native busy
host checks painted ordinary and stale-context deletions in 6 and 16 ms.

The strengthened setup changes the interval before each measured cycle, so it
is not directly comparable to the earlier permissive setup. Component bundles
and work counts establish the scan reductions; these UI captures establish
continued recovery and report remaining slow samples. They do not establish
that all intermittent editor latency is resolved.
