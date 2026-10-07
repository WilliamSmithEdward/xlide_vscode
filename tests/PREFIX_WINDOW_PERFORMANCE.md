# Completion and hover prefix windows

Typing changes the caret offset, and moving the mouse changes the hovered symbol. The four-entry full cursor-context cache cannot cover either stream reliably. Type detection, member receiver resolution and keyword classification previously copied/filtered all tokens preceding each new offset, although they consume only five grammar tokens or the current logical line.

The new helpers select their window before copying the cached token stream. Type detection retains the last five non-comment/non-newline tokens, preserving existing semantics. Member and keyword resolution keep the current logical line and its leading newline. Continuation underscores and truncated boundary tokens use the existing residue/tail re-lexing logic. The full-prefix API remains available for consumers that inspect preceding statements.

## Reproduction

Run `npx vitest run tests/completionPrefixWindows.test.ts tests/vbaCursorContext.test.ts tests/typeLookbackPerformance.test.ts`. Tests compare token kind/text/spans at every offset against fresh truncated-prefix lexing, exercise continued chains and colon statements, and count token reads on a 1,201-procedure source (fewer than 100 per paired line/type query). Lexer trivia metadata is excluded, as it is in the existing full-prefix derivation test.

Set `XLIDE_PREFIX_WINDOW_BENCH=1` and run `npx vitest run tests/completionPrefixWindows.test.ts -t "measures completion"`. The fixture has 44,768 characters and six distinct late caret positions (more than the four-entry cache). Each of 21 batches makes 60 queries. Median per-query batch times on this Windows host, compared with base 7b25e7b7:

| Warm analyzer query | Before (ms) | After (ms) |
| --- | ---: | ---: |
| Type-position check | 0.18736 | 0.00147 |
| Member completion | 0.20708 | 0.02856 |
| Keyword classification | 0.24715 | 0.00131 |
| Member hover | 0.13330 | 0.03187 |

These are component timings with unchanged, warmed source. They measure neither provider scheduling nor menu/tooltip painting, and establish no cold-start or UI tail-latency bound. First lexing of edited source still costs module-sized work. Long logical lines and long runs of trivia preceding the last five grammar tokens still require scanning those tokens. Identifier completion retains the full-prefix API because some context checks inspect earlier statements.

The real VS Code suite `Prefix window surfaces` opens a workbook module with 1,200 padding procedures, moves between six member chains, checks hover and colon-statement keywords, then edits a continued member chain one character at a time and checks completion after every edit. Existing editor, type-lookback and hover snapshot suites cover fresh-source invalidation and scope changes.

## Validation of the combined change

Compilation and the full unit suite passed before rebasing onto the merged immediate-completion scheduling change: 680 files, 13,598 passed tests and 22 intentional skips. After the rebase, 26 focused unit checks passed (3 opt-in benchmark skips), including response scheduling, prefix windows and hover snapshots; 15 real VS Code 1.139.1 harness checks passed.

The final warmed actual-typing run recorded 12 edit-to-completion-result samples: median 17.0075 ms, slowest 20.8170 ms. Existing editor-surface warm hover commands recorded a 0.8580 ms median. These are observations of the combined changes in one run, with background project facts already loaded; they are not an isolated speedup comparison for prefix windows, cold-start results, UI painting measurements or a tail-latency guarantee. Broader provider command samples remained slower (completion median 54.0321 ms), so more response-path work remains.
