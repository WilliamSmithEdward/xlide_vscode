# Move to Module qualified reference binding

Moving Reports.Build to Helpers previously missed escaped identifiers and explicit continuations, and changed unrelated object members when Reports was a parameter, local or module field or appeared in a longer receiver chain.

The scanner now uses shared lexer tokens and the existing ProjectIndex bare-identifier resolver with memberReceiver context. Only an unresolved bare receiver name denotes the standard module qualifier. It preserves bracketed receiver spelling and all surrounding trivia, including member spelling, comments, strings and line continuations. AddressOf references are repointed because the procedure signature is unchanged. One lazy query-owned project index is shared across callers and is not built without plausible candidates.

Validation against main 37b23514:

- 36 complete-output cases across LF/CRLF/CR: 30 fail before the change, six controls pass; all pass after it.
- 129 focused tests across seven files pass; type checking passes.
- Full unit suite: 716 files pass, six skipped; 14,306 tests pass, 30 skipped. No failures.
- Frozen ASTs, lexer tokens and trivia: all 16,500 full results and rendered outputs across 8,256 corpus sources match the parent; 72 generated recursive moves have independently expected complete outputs. Explicit candidate bounds/overlap checks pass; no exceptions.
- Binding work tests at 1/100/1,000 references in each of two callers require exactly one setModule call per input module and verify complete output. Comment/string/long-chain-only inputs construct no project index.

Matched benchmark, Node 24.18.0 / Ryzen 7 9800X3D, three warmups and nine measured rounds, parent/candidate/candidate/parent order. Both versions produce the expected complete primary, target and caller sources in all 18 fixtures. For 1,000 references, fresh colon medians change from 0.171–0.196 ms to 1.664–2.358 ms; fresh multiline from 0.193–0.198 ms to 1.056–1.504 ms; fresh comments from 0.340–0.346 ms to 1.833–1.892 ms. Cached results vary by layout (roughly 0.168–0.570 ms before, 0.207–0.482 ms after). This is a correctness repair, not a universal performance improvement. Reproduce with node scripts/benchmark-move-to-module.mjs --baseline=37b23514 --rounds=9 and the same command without baseline.

The Move API continues to assume standard modules. Private visibility, destination name capture and nonstandard module eligibility require separate audits; this change does not establish that all moves are semantically safe.
