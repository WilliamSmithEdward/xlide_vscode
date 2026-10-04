# Introduce Parameter: recursive binding and call syntax

Issues: #982, #1002, #1003, #1004 and #1010.

Introduce Parameter previously skipped the selected procedure's entire body. Its signature changed while recursive calls retained the old argument list. The replacement applies the existing bare-name and member-definition resolvers within that procedure, preserving function/property result reads and writes. True recursive calls pass the original initializer, preserving per-invocation initialization. Foreign receivers keep their original call arguments.

Within the selected procedure, one query-owned ProjectIndex builds each needed module at most once. The member context reuses the parsed module, significant source tokens and existing receiver, surface and With-stack caches. The command supplies module kinds through the existing host type mapping, so class/document/UserForm Me receivers can resolve. The pure API defaults omitted kinds to standard modules.

Unresolved receiver calls and self-calling initializers are refused instead of returning partial edits. Initializers reading the function's own result variable are also refused. Property Let/Set and getters with a paired setter are refused because coordinated assignment-site/accessor changes are required; unpaired getters remain supported.

The shared member resolver now recognizes bracketed member names. The call scanner retains labels, named-argument keys, AddressOf pointers, bang fields and type names instead of adding arguments to them. Numbered statements and leading-dot With calls receive proper statement context. The previous test expecting Go: to change as a call was corrected: Go: is a label. The separate external-module binding fix from merged PR #999 is retained when integrating onto main.

## Performance finding

The first complete-refactor benchmark exposed a shared member-resolver issue: Then .Report first treated Then as an explicit receiver, then scanned preceding statements looking for a Set assignment before resolving the With receiver. Deterministic lookup counts were 5,050 at 100 calls and 500,500 at 1,000 calls. The fix recognizes Then/Else/Call control-flow contexts to bypass that inappropriate explicit-root path; grouped, keyword-named host receivers and chained members retain their resolution paths. Work guards bound statement reads and module builds at 1, 100 and 1,000 calls, and verify complete edited sources with frozen ASTs.

## Measurement

Run node scripts/benchmark-call-site-arguments.mjs --recursive-refactor --rounds=9, with --baseline=f6e41f0b for the old implementation. Node v24.18.0, AMD Ryzen 7 9800X3D 8-Core Processor           ; three warmups, nine timed rounds, two independent processes per version. Baseline / first binding implementation were measured sequentially ABBA; two final candidate runs followed the With fix. Timing includes the complete pure refactor; application and full-text verification happen outside the timed interval. Fresh runs vary a leading comment, invalidating source-keyed caches; cached runs reuse source text.

All 18 final configurations produce the independently expected signature, removals and recursive arguments. All 18 old configurations omit recursive arguments. Their timings measure different work and cannot support a speedup claim. The intermediate and final implementations both produce correct output and expose the cost of the With repair. These measurements do not establish Office/UI latency or every source size's asymptotic behavior.

| Calls / receiver / source | Old (ms, wrong output) | Binding before With fix (ms) | Final binding (ms) |
|---|---:|---:|---:|
| 1/bare/cached | 0.038–0.039 | 0.082–0.088 | 0.074–0.074 |
| 1/bare/fresh | 0.107–0.116 | 0.164–0.177 | 0.136–0.150 |
| 1/Me/cached | 0.013–0.013 | 0.070–0.094 | 0.065–0.076 |
| 1/Me/fresh | 0.048–0.057 | 0.132–0.173 | 0.117–0.129 |
| 1/WithMe/cached | 0.013–0.014 | 0.104–0.115 | 0.082–0.089 |
| 1/WithMe/fresh | 0.052–0.058 | 0.143–0.154 | 0.105–0.114 |
| 100/bare/cached | 0.089–0.118 | 0.438–0.439 | 0.390–0.428 |
| 100/bare/fresh | 0.654–0.749 | 0.738–1.054 | 0.866–0.880 |
| 100/Me/cached | 0.126–0.174 | 0.601–0.627 | 0.458–0.516 |
| 100/Me/fresh | 0.535–1.079 | 0.693–1.043 | 0.903–0.908 |
| 100/WithMe/cached | 0.075–0.163 | 1.393–1.834 | 0.525–0.538 |
| 100/WithMe/fresh | 0.640–0.659 | 1.527–2.398 | 0.755–0.771 |
| 1000/bare/cached | 0.531–0.605 | 2.161–3.579 | 2.470–2.471 |
| 1000/bare/fresh | 4.507–5.571 | 6.250–8.244 | 5.786–5.827 |
| 1000/Me/cached | 0.926–0.931 | 2.939–4.267 | 2.823–2.830 |
| 1000/Me/fresh | 3.275–3.516 | 5.751–8.329 | 5.419–5.579 |
| 1000/WithMe/cached | 0.500–0.501 | 27.236–30.854 | 3.788–3.894 |
| 1000/WithMe/fresh | 2.936–3.078 | 38.680–45.147 | 6.388–7.091 |


## Validation

- Final correctness baseline: 126 failures and 15 passing controls across 141 tests; the candidate passes those cases. The work guard separately demonstrates the intermediate quadratic With scans.
- Focused checks include command module-kind propagation, original call argument fidelity and all new binding/syntax/property cases. Type checking passes.
- Differential: 8,256 corpus sources / 19,814 complete call-site queries, plus 216 generated full refactors with independently expected primary and external source. Lexer tokens/trivia and ASTs are frozen; edit bounds and overlaps are explicitly checked.
- The final differential on integrated main preserved 19,738 complete arrays and rendered sources; 76 changes were reviewed: procedure-pointer operands, labels and numbered result assignments stopped receiving spurious arguments. All 216 generated ordinary full refactors remained identical and correct. The combined focused suite passes 260 tests across 10 files; types pass. Full-suite verification on main base 975c4397 passes: 705 test files passed, two skipped; 14,152 tests passed and 26 skipped. The first full run caught four keyword-named host receiver regressions in the broad With fast path; the narrowed control-flow path passes all affected host tests.
