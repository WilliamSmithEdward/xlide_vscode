# Analyzer formatter performance — 2026-10-03

The formatter lexed the original module once to format it and again to compare tokens with the output. Reuse the unmodified original token stream in the safety check. Standalone tokenStreamDifference keeps its public source-string API and still lexes both inputs.

The continuation-placement check also built per-token arrays, filtered each trivia array and joined a string. Count leading/trailing continuation trivia directly. Token count, kinds, raw text, keyword/identifier case allowances, total continuation count and placement checks retain their order and refusal messages. No safety checks are removed.

Run `node scripts/benchmark-formatter.mjs --rounds=15` in each checkout. Node 24.18.0 / Ryzen 7 9800X3D, three warmups and 15 samples; median milliseconds. Format workloads include output validation; validation workloads compare a source with itself through the standalone API.

| Workload | Before | After |
| --- | ---: | ---: |
| Format 100 procedures | 9.011 | 4.593 |
| Validate 100 procedures | 3.506 | 1.769 |
| Format 2,000 procedures | 68.295 | 41.354 |
| Validate 2,000 procedures | 44.624 | 26.942 |
| Format 10,000 statements | 50.609 | 32.486 |
| Validate 10,000 statements | 34.933 | 19.359 |
| Format 2,000 continued statements | 9.832 | 7.315 |
| Validate 2,000 continued statements | 5.660 | 3.495 |

These are synthetic formatter measurements, not end-to-end editor latency claims. The lexer-call regression verifies one input scan and one output scan, and fails on the old implementation's three scans. Formatter unit, corpus and command tests cover spacing, casing, comments, strings, labels, indentation and continuation refusals. Full-suite validation is recorded in the PR.
