# Documentation tag source span lookup

The documentation diagnostic and parameter-reference scans call `scanDocTags` to translate offsets in a joined comment body back to physical source offsets. Each lookup previously walked backward from the last comment line. Four lookups per named self-closing tag therefore required quadratic work as the block grew.

The lookup now uses an upper-bound binary search over the already-built line-start offsets. It preserves exact-start, end-of-line and multiline-tag mapping, including directive lines omitted from the joined body. Closing spans can precede the next lookup in offset order, so a forward-only cursor would not preserve the general contract.

## Measurement

Run `node scripts/benchmark-doc-tag-spans.mjs`, or add `--baseline=e2ca4515` for the original implementation. The script measures the complete public tag scan, including construction of the joined body and line index, attribute parsing and span mapping. Fixture construction is outside the timed region. Three warmups and 15 samples, Node 24.18.0, AMD Ryzen 7 9800X3D; milliseconds:

| Self-closing tags | Original median / p95 | Optimized median / p95 |
| --- | --- | --- |
| 10 | 0.018 / 0.038 | 0.014 / 0.041 |
| 100 | 0.103 / 0.327 | 0.092 / 0.332 |
| 1,000 | 1.547 / 2.186 | 0.755 / 1.738 |
| 10,000 | 65.403 / 84.287 | 7.145 / 9.791 |

The largest fixture is a stress case, not a typical procedure comment. These figures apply to tag scanning, not complete module analysis. Searches for closing and reopening tags remain unchanged; this change does not establish a linear bound for arbitrary malformed XML bodies.

## Validation

LF and CRLF regression fixtures cover directive gaps, multiline opening tags, decoded names, repeated attributes, nested tags whose mapped offsets arrive out of order, and self-closing tags. A large block checks all 2,000 name spans. A separate differential comparison matched all returned tag fields on 500 generated documentation blocks against the original implementation.
