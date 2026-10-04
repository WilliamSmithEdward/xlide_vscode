# Refactor edit chunk assembly

The shared applyVbaTextEdits helper rebuilt the full source after every edit. It is used directly by multi-module refactor commands, inside Extract Method for grouped-declaration changes and by Move to Module for moved-body rewrites. A 100,000-character source with 1,000 replacements produced 99,999,000 returned slice characters, even though the final output remained source-sized.

Keep stable descending edit order. Validate that all spans are integer, in range and nonoverlapping in that order; collect untouched source regions once and join the chunks. Equal-offset insertions retain the original order of application. Overlapping, reversed, fractional, negative, infinite, NaN and out-of-range spans retain the existing sequential String.slice path, preserving the public helper's behavior. Single-edit calls also retain sequential construction. Caller arrays/spans are never mutated. Temporary chunk storage scales with edit count; no retained cache is added.

Nineteen tests cover deterministic source-sized slice work, stable same-offset insertion/replacement order, overlap, unusual offsets, Unicode UTF-16 text, newline forms and frozen inputs. Two work checks fail on baseline; seventeen compatibility controls already pass, including 2,000 generated edit sets. All nineteen pass after restoring exact production bytes. Types, 104 focused checks and the full suite pass: 14,977 tests, 33 skipped, 767 passing files, seven skipped files, 85.49 seconds.

All 74,304 complete comparisons match baseline across 8,256 oracle sources and three newline styles: 24,768 edit applications, 24,768 complete Extract Method results and 24,768 complete Move to Module results. All edit/extract queries succeed; 24,459 move queries succeed and the remaining refusal results also match. Corpus refactor queries include independent injected grouped locals/qualified recursive calls. Independent benchmark outputs and explicit compatibility tests supplement the differential coverage. No arbitrary malformed-module compilation claim.

## Reproduction and timing

Run `node scripts/benchmark-refactor-edit-chunks.mjs` in baseline/current/current/baseline order, supplying `--baseline=c9227136fa687247d749c84e973a90f6e80d0193` for baseline runs. Node 24.18.0, Ryzen 7 9800X3D; three warmups, nine measured rounds, one complete call per round. Source/edit setup is outside the helper clock; source/AST warming is outside the Extract Method clock. Complete output is independently constructed and checked afterward.

The 10,000-edit case uses 640,000 source characters and improves from 1153.2112–1156.8608 ms to 0.7281–0.7389 ms. Actual complete Extract Method with 1,000 grouped declarations improves from 14.6711–15.2228 ms to 11.7930–12.6228 ms. Single-edit controls are slightly slower, from 0.0010–0.0012 to 0.0013–0.0016 ms; single-declaration extraction is mixed. All baseline/current outputs agree. No cold-parser, heap-byte or editor-command/renderer-paint improvement is claimed.

Ranges below are medians from two independent runs, milliseconds per complete warmed call.

| Query | Count | Baseline ms | Repair ms |
| --- | --- | --- | --- |
| applyEdits | 1 | 0.0010–0.0012 | 0.0013–0.0016 |
| applyEdits | 1000 | 1.2353–1.2716 | 0.0970–0.1098 |
| applyEdits | 10000 | 1153.2112–1156.8608 | 0.7281–0.7389 |
| extractMethod | 1 | 0.0380–0.0457 | 0.0345–0.0423 |
| extractMethod | 100 | 0.6516–0.7291 | 0.4800–0.4813 |
| extractMethod | 1000 | 14.6711–15.2228 | 11.7930–12.6228 |
