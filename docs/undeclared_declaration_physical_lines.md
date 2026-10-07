# Undeclared declaration physical lines

The undeclared-variable declaration fix used LF-only line helpers. In CR-only procedures it inserted before the module at offset zero and lost indentation, with or without preceding declarations. Its first-body lookup also skipped block statements: a continued procedure header followed by a block could receive a declaration inside its parameter list even in LF/CRLF text.

Use shared nearest-break start/end scanners for insertion and indentation, and step over CRLF as one terminator when finding a following line. Let the first-body lookup recognize block nodes as well as leaf statements. Retain the existing one-time insertion-site setup per procedure and lazy once-per-module EOL lookup. Shared CR serialization comes from PR #1132; no persistent cache is added.

Thirty independent complete-fix/applied-output cases cover LF/CRLF/CR, simple and continued headers, preceding declarations/comments, block bodies and 0/1,000 prefix lines. Fourteen fail on the old undeclared rule, including continued-header/block placement in LF/CRLF; sixteen already pass. All thirty pass after restoring exact production bytes. Types and 62 focused checks pass. The full combined suite, including PR #1132's production changes and tests, passes 14,988 tests, 33 skipped, 767 passing files, seven skipped files, 84.10 seconds.

All 16,512 complete diagnostic/fix results match baseline for LF/CRLF, and 8,256 current CR results match canonical current LF after newline/span normalization. Each of the 8,256 unfiltered oracle sources has an injected independent missing assignment, so all 24,768 queries include a declaration fix. New independent expectations establish the header/block corrections missing from the corpus. This is rule-level coverage, not arbitrary malformed-module compilation in Office.

## Measurement

Run `node scripts/benchmark-undeclared-physical-lines.mjs` in baseline/current/current/baseline order with `--baseline=275adc05ba103a2f0363ca83e6889301b3153bda` on baseline runs. The baseline override replaces only the undeclared rule and keeps PR #1132's CR EOL helper, isolating insertion/indent lookup. Node 24.18.0, Ryzen 7 9800X3D; three warmups, nine measured rounds, ten complete rule queries per round.

Source, AST and symbols are outside the clock. Complete independent fix data, finding count/rule and applied insertion output are checked afterward. CR baseline intentionally checks the old wrong insertion, so CR timings are not equal-output comparisons. With 10,000 prefix lines, complete warmed CR rule calls improve from 0.08408–0.08838 ms to 0.04360–0.04755 ms. LF/CRLF and tiny controls are mixed. No full-analyzer, cold-parser, heap-byte or editor-latency gain is claimed.

Ranges below are medians from two independent runs, milliseconds per complete warmed rule query.

| Prefix lines | EOL | Baseline ms | Repair ms |
| --- | --- | --- | --- |
| 0 | "\n" | 0.00991–0.01259 | 0.01029–0.01065 |
| 0 | "\r\n" | 0.00872–0.01089 | 0.00925–0.01033 |
| 0 | "\r" | 0.00652–0.00776 | 0.00722–0.00753 |
| 10000 | "\n" | 0.01023–0.01205 | 0.01137–0.01599 |
| 10000 | "\r\n" | 0.00441–0.00553 | 0.00543–0.00721 |
| 10000 | "\r" | 0.08408–0.08838 | 0.04360–0.04755 |
