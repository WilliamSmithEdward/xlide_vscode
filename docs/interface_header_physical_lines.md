# Interface header extraction and duplicate scanner removal

Implement Interface had a private LF-only header scanner. For CR-only interfaces it copied a procedure body and closing statement into the generated header, then appended another body and closer. Mixed line endings could truncate a continued parameter list. These results violate the refactor's promise to copy the exact interface signature.

The refactor now uses `blockHeaderLineSpan`, the existing CR/LF/CRLF-aware scanner. Its implementation and first-break helper move verbatim from diagnostics/walker.ts to parser/physicalLineSpans.ts, with the old walker exports preserved. This lightweight module avoids pulling diagnostic/dataflow dependencies into the interface refactor. The duplicate private header scanner and its LF-only dependency are removed. Existing continuation recognition requires the VBA space/tab before `_`; malformed trailing identifier underscores are no longer mistaken for continuation markers.

Eleven new exact-edit tests cover body exclusion at 1/1,000 lines, all three physical line endings, continued signatures with mixed endings, ByRef/Optional defaults, modifiers and trailing comments. Four cases fail on baseline; 84 focused checks pass, including interface refactoring, Introduce Parameter and diagnostic block-header consumers. The helper extraction is verified against the exact baseline text after newline normalization. Types pass.

## Measurement

Run `node scripts/benchmark-interface-header-lines.mjs` in baseline/current/current/baseline order, supplying `--baseline=6b097a1eca852560946804a2299f49fd10ae0a36` on baseline runs. Node 24.18.0, AMD Ryzen 7 9800X3D. Three warmups, nine measured rounds, twenty complete refactor calls per round. Both class and interface ASTs are warmed before the clock; source construction and result checking are excluded. Baseline CR output is independently checked as the known corrupted result, so this is a correction rather than an output-parity speedup.

For a 10,000-line interface body:

| Interface line endings | Baseline median ms | Repair median ms | Generated edit characters before → after |
| --- | --- | --- | --- |
| LF | 0.00264–0.00329 | 0.00250–0.00256 | 107 → 107 |
| CRLF | 0.00243–0.00252 | 0.00211–0.00335 | 107 → 107 |
| CR | 0.05229–0.05507 | 0.00243–0.00302 | 310,115 → 107 |

Ranges are the medians of two independent runs. Maximum CR batch-average samples are 0.05796–0.06722 ms before and 0.00540–0.00643 ms after. LF/CRLF controls are mixed and tiny fixtures show ordinary variation. Payload characters are measured, not heap allocation bytes. This does not measure cold parsing, editor command latency, renderer painting or actual Office compilation of arbitrary mixed-ending input. The implementing class retains the existing generated-line-ending policy; original continued signature text is preserved.

## Complete validation

Full suite: 14,751 tests pass, 33 skipped, 759 passing files and seven skipped files (86.68 s). All 17,112 LF/CRLF complete refactor comparisons match the baseline over 8,556 corpus/malformed sources. Another 17,112 CR/CRLF complete results match the repaired canonical LF result after normalizing edit text line endings; class edit spans and titles/refusals remain exact. The existing scanner moved verbatim, verified against baseline source text. Corpus parity does not establish Office compilation of every malformed interface fixture.
