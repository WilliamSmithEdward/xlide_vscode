# Stack-safe constant dependency resolution

`resolveRawIntegerConstants` previously evaluated an unresolved dependency by recursively entering another expression parser. Expression nesting guards were local to each parser, so a long forward dependency chain could overflow the JavaScript stack. A 10,000-constant forward arithmetic chain and a 10,000-constant cycle throw `RangeError`; reverse declaration order resolves the same acyclic values. Constants wrapped in nested parentheses can overflow with much shorter dependency chains.

The shared expression parser now yields constant names and resumes with their values. An explicit stack schedules unresolved dependencies and resumes the suspended parser directly. Its cursor and partial arithmetic values survive suspension: wide expressions do not replay already-read prefixes. The standalone evaluator drives the same parser with its existing lookup, preserving one grammar. Single literals and aliases bypass precedence frames. Cycles, ambiguous duplicates, early unknowns, base lookups and result Map order retain their behavior. Existing per-expression nesting limits remain.

## Validation

Ten new regressions cover 10,000 aliases, arithmetic dependencies and cycles; a 10,000-dependency wide expression; combined expression/dependency nesting; complete ordered maps and lookup order; token paths for Not; and both diagnostic collection and project exports across LF/CRLF/CR. Seven of the ten new cases fail with stack overflow on the baseline; three controls pass. Type checking and 31 focused tests passed.

Independent comparison against `d02b4190b7afa65a5be38e3f210c0baa8df6f6b0` matched 9,996 expression values and lookup orders (including malformed text, numeric boundaries, calls, and nesting limits), 3,000 complete ordered generated constant maps, and 8,256 complete analyzer diagnostic/failure results.

## Measurement and tradeoff

Run `node scripts/benchmark-constant-dependencies.mjs`, optionally with `--baseline=d02b4190b7afa65a5be38e3f210c0baa8df6f6b0`. Inputs and complete independent expected-result assertions are outside the timer. Evaluation includes tokenization, parser allocation and lookups, with lexer caches warmed. Separate before/after/after/before processes use three warmups and nine rounds. Environment: Node 24.18.0, AMD Ryzen 7 9800X3D.

| Workload | Before median (ms) | After median (ms) |
| --- | --- | --- |
| 10,000 forward arithmetic constants | Stack overflow in all nine samples per run | 14.6314–15.2061 |
| 10,000 cyclic arithmetic constants | Stack overflow in all nine samples per run | 12.6237–12.6823 |
| 10,000 reverse-ordered arithmetic constants | 7.3230–7.5158 | 12.0739–12.0842 |
| 10,000 dependencies in one sum | 7.5240–7.5310 | 9.5608–9.5959 |
| 100 forward arithmetic constants | 0.10566–0.10580 | 0.17555–0.18003 |
| Single integer literal | 0.00017–0.00019 | 0.00012–0.00013 |
| Single alias | 0.00019–0.00020 | 0.00013–0.00014 |
| `name + 1` | 0.00059–0.00061 | 0.00108–0.00110 |
| `CInt(3.5) + (name * 2)` | 0.00159–0.00160 | 0.00268–0.00277 |

Resumable precedence frames have a measurable throughput cost. This repair prevents dependency stack overflow; it does not claim a universal speedup. Suspended state uses heap space proportional to active dependencies and expression state; heap bytes and editor latency were not measured. Failed baseline samples are not comparable completed-work timings.

Final full suite: 752 files passed, seven skipped; 14,682 tests passed, 33 skipped (79.20 seconds). The external-lookup no-replay assertion was strengthened afterward and all 31 focused tests passed again; production code was unchanged.

A follow-up complete-analyzer control measured warmed `analyzeModule` with default options, alternating before/after order, three warmups and nine rounds; independent complete diagnostics and zero internal failures were checked outside the clock. Medians: 100 forward constants, 2.6690 -> 3.3993 ms; 1,000 reverse-ordered constants, 15.1905 -> 20.2834 ms; 1,000 functions returning `2 + 3`, 129.5127 -> 126.4348 ms. Maxima were 6.0647 -> 4.3384 ms, 17.0482 -> 20.4515 ms, and 215.9428 -> 165.0415 ms respectively. These synthetic controls confirm the throughput cost on constant-heavy modules; they do not establish a live-editor speedup.
