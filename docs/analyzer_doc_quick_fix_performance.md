# Documentation quick-fix line lookup

Documentation diagnostics create removal fixes for duplicate tags. Every fix looked up its starting and ending physical documentation lines by scanning every line in the block. Repeated tags made this lookup quadratic even after tag scanning had been optimized.

`DocBlock.lineAt` now reuses the shared `firstTokenAtOrAfter` lower-bound helper on the already ordered documentation line starts, then chooses the exact or preceding line. The first-line clamp, directive gaps, CRLF offsets, whole-line removal and shared-line tag removal retain their existing behavior. No new index or persistent cache is allocated.

## Measurement

Run `node scripts/benchmark-doc-comment-fixes.mjs`, or add `--baseline=661c242e` for the original diagnostic rule. The benchmark times the complete `checkDocComments` call on a procedure with repeated self-closing summaries, including tag scanning and quick-fix allocation. Source construction and module parsing are excluded. Three warmups and 15 samples, Node 24.18.0, AMD Ryzen 7 9800X3D. Milliseconds, median / p95:

| Duplicate-summary fixture | Original | Optimized |
| --- | --- | --- |
| 100-duplicate-summary-fixes | 0.141 / 0.425 | 0.129 / 0.228 |
| 1000-duplicate-summary-fixes | 1.837 / 2.386 | 0.615 / 1.142 |
| 10000-duplicate-summary-fixes | 181.043 / 213.92 | 5.738 / 7.325 |

The largest case is a stress block producing 9,999 duplicate-tag findings. These measurements apply to this rule and its edit generation, not an end-to-end module analysis or typical short documentation comment.

## Validation

A structural regression wraps physical line-start reads. The original full scans exceed the linear read budget for 200 tags; binary lookup meets it while every removal edit still targets the exact expected physical line. LF/CRLF fixtures verify removal of attached next-line directives, and a shared-line fixture verifies that only the repeated tag is removed. Complete diagnostic arguments, including all quick-fix objects and edit spans, match the original on 1,500 generated comments.
