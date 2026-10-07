# Extraction physical lines and CR serialization

Extract Variable used an LF-only line-start lookup. In CR-only modules it inserted the declaration before the procedure, lost indentation and generated LF text. Mixed modules containing a lone CR between statements could place the declaration before the wrong statement. The shared serialization EOL helper also returned LF for CR-only text, so generated interface/property text and quick fixes mixed newline styles.

Use the existing nearest-break line-start scanner for extraction. Extend the shared detectEol helper to return CR for CR-only text, retaining CRLF precedence and existing LF/mixed/default behavior. Remove three now-redundant CR fallbacks from Extract Method, Move to Module and documentation edits. Two existing tests explicitly encoded CR-to-LF conversion; update only their EOL expectations, retaining complete edit and work-count assertions.

Twenty-eight independent new cases cover EOL categories, simple/nested indentation, 0/1,000 preceding lines, mixed boundaries, exact applied extraction output and rename spans, and generated interface/property output. Nine fail on baseline; nineteen already pass. All 28 pass after restoring the exact production bytes. Types, 198 focused checks and the full suite pass: 14,922 tests, 33 skipped, 765 passing files, seven skipped files, 73.53 seconds.

Complete-result differential coverage uses all 8,256 oracle sources. Five query families each have 16,512 exact LF/CRLF comparisons and 8,256 canonical CR comparisons: extraction from derived parsed bodies, interface stubs, field encapsulation and Dim-initializer/Option Explicit quick fixes. Total: 82,560 baseline-compatible results and 41,280 CR results matching current LF after converting newline text and span coordinates; zero differences. Rename coordinates are normalized against the applied output. This covers the exercised queries, not arbitrary Office compilation or every quick-fix placement path.

## Measurement

Run `node scripts/benchmark-extract-physical-lines.mjs` in baseline/current/current/baseline order, supplying `--baseline=3f394552e3496dae695b03e19529f28d338d2bd5` on baseline runs. Node 24.18.0, Ryzen 7 9800X3D; three warmups, nine measured rounds, ten complete queries per round. Supply an explicit extraction name to isolate line/EOL work from generated-name allocation. Source construction and AST warming are outside the clock; complete output/title/insertion/rename controls are checked afterward.

CR baseline outputs deliberately check the known incorrect insertion at module start and LF text, so CR timings are not equal-output comparisons. With 10,000 prefix lines, complete CR queries improve from 0.15157–0.15532 ms to 0.10100–0.10442 ms. Tiny CR costs are mixed/slower and LF/CRLF controls are mixed. No cold-parser, heap-byte, editor-command or renderer-paint claim is made.

Ranges are medians from two independent runs, milliseconds per complete warmed query.

| Prefix lines | EOL | Baseline ms | Repair ms |
| --- | --- | --- | --- |
| 0 | "\n" | 0.01208–0.01353 | 0.01259–0.01326 |
| 0 | "\r\n" | 0.00689–0.00708 | 0.00684–0.00708 |
| 0 | "\r" | 0.00552–0.00628 | 0.00623–0.00710 |
| 10000 | "\n" | 0.06745–0.07891 | 0.06542–0.06868 |
| 10000 | "\r\n" | 0.05985–0.05997 | 0.05959–0.06283 |
| 10000 | "\r" | 0.15157–0.15532 | 0.10100–0.10442 |
