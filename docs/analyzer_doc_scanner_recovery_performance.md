# Documentation attribute and opening scanner recovery

Two remaining documentation scanner paths retried malformed text quadratically. Attribute matching retried each suffix of a long unknown identifier without an assignment, and diagnostic opening-tag matching retried the remainder of a block for each incomplete tag prefix when no `>` remained.

A shared attribute iterator now matches names independently, checks the following assignment/opening quote at that exact position, and consumes the first closing quote before the next name. It preserves the existing ASCII name grammar, matching after invalid leading characters, duplicate last-attribute-wins behavior, entity decoding and raw value spans. Both `attrsOf` and `scanDocTags` use this iterator; the duplicate regex parsing loops and unused attribute regex are removed.

Diagnostic tags now use the existing vocabulary-prefix regex and advance to their first `>`, stopping when no closing angle bracket remains. They retain the earliest opening, lenient malformed attribute text, self-closing detection and physical source mapping. The separate reopening/paired-close interpretation is unchanged.

## Measurement

Run `node scripts/benchmark-doc-scanner-recovery.mjs`, or add `--baseline=f7fa3f54` for the original. Three warmups and 15 samples, Node 24.18.0, AMD Ryzen 7 9800X3D. Unknown-word fixtures time the complete `parseDocBody` call; incomplete-tag fixtures time the complete `scanDocTags` call. Ordinary controls time 1,000 model parses and corresponding tag scans. Fixture construction excluded. Milliseconds, median / p95:

| Fixture | Original | Optimized |
| --- | --- | --- |
| 100-unrecognized-attribute-word | 0.013 / 0.025 | 0.005 / 0.018 |
| 100-incomplete-diagnostic-openings | 0.015 / 0.036 | 0.004 / 0.015 |
| 1000-unrecognized-attribute-word | 0.425 / 0.511 | 0.003 / 0.004 |
| 1000-incomplete-diagnostic-openings | 0.836 / 1.53 | 0.044 / 0.344 |
| 10000-unrecognized-attribute-word | 38.034 / 39.795 | 0.009 / 0.024 |
| 10000-incomplete-diagnostic-openings | 85.531 / 99.438 | 0.181 / 0.413 |
| 1000-ordinary-models-and-tag-scans | 7.413 / 9.59 | 7.02 / 9.564 |

The largest fixtures are malformed editor-content stress cases. These figures apply to the measured documentation functions, not complete analysis, hover rendering or typical short comments. A repeat measurement was used because the initial run had greater timing variance.

## Validation

All returned documentation models and diagnostic tag fields, including spans, match the original on 27,000 three-fragment combinations and 10,000 seeded generated bodies. Regression fixtures cover large malformed inputs, invalid name prefixes, failed unquoted values, repeated attributes, decoded names with raw entity spans, LF/CRLF, valid tags before incomplete suffixes, nested opening recovery and valid hints before unterminated attributes.
