# Worksheet-change statement scan

The conservative worksheet-change scanner repeated a whole-statement range-word search for each Copy member. Name comparisons also copied every remaining token before reading only the first two RHS tokens. Keep the range-word result lazily within its statement and inspect the RHS by token index. Existing decisions and ordered name results remain unchanged; no new persistent cache is added.

Validation against 4d66f518:

- Full suite with two workers: 721 files passed, seven skipped; 14,368 tests passed, 31 skipped. No failures. The first default-parallel run had 14,366 passes and two unchanged-test failures: inflate exceeded its 5s timeout; undeclaredVariableCost measured 21.298 against its <20 ratio. Both files passed unchanged in isolation (22 tests), then the complete two-worker suite passed. No thresholds, timeouts or exclusions changed.
- Six work-count cases: four baseline failures and two one-member controls. At 1,000 Name comparisons, 3,003,000 copied token entries fall to zero. Range-word token reads are bounded linearly in statement size for 1/100/1,000 Copy members.
- Frozen-input differential: all 8,328 complete sheet-fact results, ordered names and diagnostic/error arrays match 4d66f518 across 8,256 corpus sources and 72 generated cases (LF/CRLF/CR, Copy/range veto/Name/renaming/comments). 13,551 findings; zero internal errors. AST nodes and lexer tokens/trivia are frozen.
- Types and 42 focused checks pass across statement work, workbook sheet diagnostics and workbook sheet metadata.

Reproduce: node scripts/benchmark-sheet-change-scan.mjs --baseline=4d66f518 --rounds=9 and without baseline. Three warmups, nine rounds; baseline/candidate/candidate/baseline order; Node 24.18.0 / Ryzen 7 9800X3D. Counts 1/100/1,000, Copy/Copy-with-range/Name-comparison forms and cached/fresh source make 18 configurations, each with independently expected complete facts.

At 1,000 members:

| Form | Before cached/fresh median ms | After cached/fresh median ms |
| --- | --- | --- |
| Copy | 50.637–55.664 | 0.297–0.568 |
| Name comparison | 0.461–1.212 | 0.095–0.345 |
| Copy with early range word | 0.105–0.290 | 0.143–0.429 |

The early-range control is already cheap and varies between runs. This removes quadratic work, not a universal speedup or total-editor latency claim. The scanner's existing conservative interpretation remains; binding/syntax improvements require separate audits.
