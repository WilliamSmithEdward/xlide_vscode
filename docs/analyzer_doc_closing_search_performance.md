# Documentation closing-tag search

The documentation diagnostic and parameter-reference scanner repeatedly searched the remaining comment body for each opening tag's closing tag. Repeated unclosed tags, or many opening tags sharing a single final close, made this search quadratic.

A scan-local map remembers the next closing offset for each of the six vocabulary tags. Opening offsets arrive in ascending order, so the remembered close is valid until the scan passes it. Once no closing tag remains, that result stays valid for the rest of this scan. Missing closes also skip the irrelevant reopening search. Reopening detection retains the existing lenient regex semantics, including markers inside another tag's attributes. The cache does not outlive one call.

## Measurement

Run `node scripts/benchmark-doc-tag-closing.mjs`, or add `--baseline=0ef9ac84` for the preceding implementation. This baseline already includes the separate binary source-span lookup from PR #748, so these measurements isolate the closing-search change. Three warmups, 15 samples, Node 24.18.0, AMD Ryzen 7 9800X3D. Complete public `scanDocTags` call, including joined-body construction, line indexing, tag parsing and span mapping; fixture construction excluded. Times are milliseconds, median / p95:

| Fixture | Before | After |
| --- | --- | --- |
| 100-unclosed | 0.152 / 0.48 | 0.114 / 0.452 |
| 100-final-close | 0.141 / 0.219 | 0.11 / 0.246 |
| 100-paired | 0.093 / 0.183 | 0.099 / 0.193 |
| 1000-unclosed | 4.604 / 5.034 | 0.771 / 1.194 |
| 1000-final-close | 4.496 / 4.963 | 1.238 / 1.805 |
| 1000-paired | 1.207 / 1.761 | 1.178 / 1.739 |
| 10000-unclosed | 491.565 / 495.949 | 9.045 / 14.251 |
| 10000-final-close | 491.622 / 495.726 | 8.873 / 12.308 |
| 10000-paired | 13.359 / 16.415 | 11.543 / 14.285 |

The 10,000-tag fixtures are stress cases, not typical procedure documentation. Paired-tag controls remain comparable. These results do not establish an end-to-end analyzer improvement or a general bound for XML parsing: `parseDocBody` and the external metadata parser are separate paths and remain under audit.

## Validation

Two structural regression tests count the characters offered to closing-tag searches. At 1,000 tags the original offers about 12 million characters; the new implementation stays below twice the source length, for both missing and final closes. Further tests cover advancing past an old close, independent tag types, self-closing tags, CRLF, and reopening markers inside attribute text. Final results match the original exactly on 1,500 generated malformed, mixed-case and Unicode documentation blocks.
