# External documentation member recovery

The external metadata entry regex retried the remaining XML body for each opening member when no closing member tag existed. Repeated unclosed entries therefore caused quadratic recovery before documentation bodies could be parsed.

The parser now recognizes the same opening syntax, searches for its first following closing member, and consumes that pair before the next opening. When no closing member remains, no later opening can make a pair, so scanning ends. Empty names still consume their complete pair before being skipped. Names retain their original trimming and entity handling. Scan positions use the original string, preserving Unicode coordinates.

## Measurement

Run `node scripts/benchmark-external-doc-members.mjs`, or add `--baseline=661c242e` for the original. Three warmups and 15 samples, Node 24.18.0, AMD Ryzen 7 9800X3D. Complete public `parseMetadataFile` calls timed, including parsing returned documentation bodies; fixture construction excluded. Milliseconds, median / p95:

| Members | Original | Optimized |
| --- | --- | --- |
| 100-unclosed | 0.029 / 0.05 | 0.001 / 0.008 |
| 100-paired | 0.038 / 0.057 | 0.043 / 0.088 |
| 100-empty-name | 0.004 / 0.215 | 0.005 / 0.012 |
| 1000-unclosed | 2.678 / 3.479 | 0.01 / 0.023 |
| 1000-paired | 0.165 / 0.267 | 0.224 / 0.471 |
| 1000-empty-name | 0.033 / 0.062 | 0.05 / 0.111 |
| 10000-unclosed | 353.243 / 490.822 | 0.062 / 0.062 |
| 10000-paired | 2.054 / 4.434 | 2.082 / 2.949 |
| 10000-empty-name | 0.585 / 0.708 | 0.405 / 0.472 |

The 10,000-unclosed-member fixture stresses malformed metadata and is not a typical documentation file. Paired and empty-name controls retain comparable behavior. This bounds the repeated member-body closing searches, not arbitrary XML processing or end-to-end analyzer latency; documentation body parsing is independently covered by PR #752.

## Validation

Regression fixtures cover an unclosed suffix, first-close pairing with nested openings, empty-name pairs consuming nested entries, Unicode text, tag casing, whitespace, literal name entities, accepted opening syntax and retention of earlier parsed entries. A separate comparison checks complete returned metadata entries against the original on 23,824 generated malformed and nested bodies.
