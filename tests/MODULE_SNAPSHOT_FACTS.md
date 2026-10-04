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
