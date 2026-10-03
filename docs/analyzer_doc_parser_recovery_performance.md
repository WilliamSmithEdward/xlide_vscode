# Documentation body parser recovery

`parseDocBody` is used by inline symbol documentation and external metadata. Its paired/self-closing regexes retried the remaining body for each opening tag when no closing tag existed. Repeated incomplete opening prefixes without a final `>` also caused quadratic recovery.

A shared tag-match iterator now locates opening prefixes, consumes their first `>`, and searches for a paired close only until it establishes that no close remains. Completed pairs consume their bodies before the next match, preserving the old non-overlapping matching behavior. Self-closing entries yield an empty body. The iterator shares the matching logic formerly duplicated between `firstTagMatch` and `extractParams` while retaining their different attribute handling. All scan state belongs to one iterator; there is no persistent source cache.

The delimiter searches advance through the body, and each missing delimiter terminates further searches for that delimiter. Attribute extraction, entity decoding and prose formatting retain their existing implementations. This is deliberately the same lenient documentation grammar, including openings embedded in malformed attribute text and a paired body that contains another opening tag.

## Measurement

Run `node scripts/benchmark-doc-body-parsing.mjs`, or add `--baseline=661c242e` for the preceding implementation. Three warmups and 15 samples, Node 24.18.0, AMD Ryzen 7 9800X3D. Complete public `parseDocBody` calls timed; fixture construction excluded. Milliseconds, median / p95:

| Fixture | Before | After |
| --- | --- | --- |
| 100-unclosed-params | 0.029 / 0.054 | 0.011 / 0.026 |
| 100-incomplete-params | 0.022 / 0.035 | 0.002 / 0.003 |
| 100-unclosed-summaries | 0.018 / 0.02 | 0.007 / 0.018 |
| 100-paired-params | 0.072 / 0.273 | 0.093 / 0.252 |
| 1000-unclosed-params | 2.396 / 3.789 | 0.044 / 0.138 |
| 1000-incomplete-params | 1.854 / 3.09 | 0.008 / 0.019 |
| 1000-unclosed-summaries | 1.42 / 1.754 | 0.041 / 0.146 |
| 1000-paired-params | 0.532 / 0.699 | 0.581 / 0.633 |
| 10000-unclosed-params | 255.745 / 275.845 | 0.411 / 0.499 |
| 10000-incomplete-params | 196.157 / 244.772 | 0.064 / 0.091 |
| 10000-unclosed-summaries | 163.215 / 194.711 | 0.385 / 0.394 |
| 10000-paired-params | 9.26 / 12.69 | 6.28 / 7.294 |
| 1000-ordinary-docs | 5.392 / 7.482 | 4.705 / 5.048 |

The 10,000-tag recovery cases are stress fixtures for malformed editor content. The paired controls parse the corresponding number of completed parameters. The ordinary-document control parses 1,000 copies of a short completed comment containing a summary, two parameters, returns, remarks and an example. These figures apply to documentation body parsing, not an entire hover, symbol build or analyzer pass.

## Validation

New recovery fixtures cover missing paired closes, incomplete prefixes, nested parameters and summaries, last-attribute-wins behavior, slash trimming, lenient tag boundaries, Unicode casing and preserved example layout through external metadata. A differential comparison matched all returned documentation fields on 27,000 three-fragment combinations and 10,000 seeded generated bodies against the original parser.
