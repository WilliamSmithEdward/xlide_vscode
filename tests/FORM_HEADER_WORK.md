# Bounded designer-header queries

## Finding and change

A retained large-class CPU capture attributed 2.50 seconds of self time to
`parseUserFormControls` over 813 recovery cycles. Both that query and
`hasAuthoritativeDesignerHeader` split the whole source into physical lines
before checking whether its first nonblank line opens with VERSION. The project
index deliberately accepts any module source here, including mislabeled forms.

The two queries now share a lazy physical-line iterator. It rejects ordinary
code at the first nonwhitespace character and yields only designer lines after
the opening VERSION line. Existing block, property, binary OleObjectBlob and
malformed-header stopping rules are unchanged. A completed designer stops before
its code body. No cache retains module source or previous revisions.

## Work-count and compatibility evidence

A 143,906-character ordinary code fixture previously allocated 20,004 line
entries across two full-source splits. A 143,966-character fixture with a small
designer header previously allocated 20,012 entries. The same two API calls now
perform zero full-source splits in both cases. These are allocation work counts,
not elapsed-time or heap-size estimates.

The regression tests cover LF and CRLF, BOM and blank prefixes, a final line
without a newline, bare CR, missing or later VERSION, binary form headers and
an unterminated header followed by code. An additional local differential check
compared 240 prefix/header/line-ending combinations with a frozen copy of the
previous parser; all agreed. The promoted regression and existing form suites
passed 137 tests across five files. Compilation and the full unit suite passed:
14,906 tests across 765 passing files, with 33 tests and seven files skipped.
The full run used four workers.

## Controlled component comparison

A local benchmark bundled the previous parser and the changed parser with the
same dependencies. Each query pair invokes both public APIs. It used 30 warmup
pairs, 100 warm pairs and 40 fresh-source pairs per variant. Fresh sources change
a fixed-width synthetic trailing comment. The private class contains 969,328
characters; a second fixture adds a public 76-character designer header to it.
Private source, bundles and detailed logs remain in ignored .vscode-test.

| Fixture and input | Previous median/p95/max (ms) | Changed median/p95/max (ms) |
| --- | --- | --- |
| Ordinary class, warm | 1.1259 / 1.8639 / 2.2973 | 0.0003 / 0.0004 / 0.0080 |
| Ordinary class, fresh | 1.4621 / 2.5028 / 3.1117 | 0.1072 / 0.7206 / 0.9547 |
| Designer plus class body, warm | 1.1846 / 2.2342 / 3.3054 | 0.0015 / 0.0022 / 0.0141 |
| Designer plus class body, fresh | 1.4549 / 2.7314 / 3.2372 | 0.1069 / 0.1500 / 1.2116 |

The work-count tests provide stronger evidence than the very small warm timings.
Fresh concatenated strings can still incur flattening costs. These component
measurements do not establish popup painting latency or resolution of every
intermittent editor stall.

## Editor validation

The pinned VS Code 1.139.1 completion surface passed all 25 cases. The private
large-class/native Backspace run passed all 12 cases, including 64 fresh-source
miss/recovery cycles and four mouse hovers. Median/p95/max observations were
Backspace 16/32/32 ms, menu recovery 118/135/137 ms, typing 31/31/32 ms and
miss clearing 1/1/11 ms. The four hovers took 376, 369, 352 and 356 ms, including
VS Code's normal hover delay. Ordinary and stale-context Backspace both painted
during an intentionally busy extension host in 16 ms.

These are continued-recovery checks, not an isolated before/after UI comparison.
The component benchmark and allocation tests establish the avoided work.
