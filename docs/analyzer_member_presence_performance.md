# Known public member presence performance

The member-not-found rule previously discarded non-exhaustive surfaces and then resolved private ownership even for known public members. It now retains a lightweight membership predicate and exhaustiveness together. Private ownership is skipped only when the known public surface contains the requested name; missing/private/unknown receivers retain the existing resolution. The exhaustive-only API retains its original object shape and absence behavior. No private-name index, persistent cache, dependency or UI change is introduced.

## Reproduce

Run from the repository root with installed pinned dependencies:

```powershell
node scripts/benchmark-analyzer-member-presence.mjs --baseline=328f46b0 > before.json
node scripts/benchmark-analyzer-member-presence.mjs > after.json
```

These measurements use Node v24.18.0 on Windows with a Ryzen 7 9800X3D. Four isolated processes ran sequentially in baseline/fixed/fixed/baseline order (A1/B1/B2/A2). Each process collected 15 timed rounds after three warm-ups. No full suite or corpus job overlapped the timings.

## Workload and limits

- The actual diagnostic rule runs through the real statement walker; `full` runs complete `analyzeModule`; `full-fresh` changes a trailing comment to bypass the parser/source cache. This is in-memory analyzer timing, with no Office execution or editor/UI timing.
- Each round gets a fresh project metadata list. The ordinary fixtures contain ten unrelated document types with ten unrelated private names each, plus C. `no-private` omits private names. Positive cases place C with private Hidden first or last; negative statements read declared Range.Value. The large cold negative contains 1,000 unrelated document types with 100 private names each and one Range.Value reference.
- Metadata generation, initial parsing/token setup for the isolated rule, complete-result assertions, and hashing are outside the timer. Per-pass receiver/member caches are fresh in the rule fixture. Whole-analyzer setup stays inside its timer.
- Every measured round matches its fixture’s complete result. Complete result hashes agree across all four processes for every row.
- The work-counter regression removes 1,000 private-owner resolver calls from 1,000 valid host references. Private document/form ownership, exhaustive-class errors, public shadowing, ambiguity, unknown receivers, mutable metadata between invocations and immutable public metadata remain covered.

## Results

All times are milliseconds. Each range spans the two process medians (or p95s) for that implementation, rather than merging samples from different processes.

| Metadata | References | Scope | Baseline median range | Fixed median range | Baseline p95 range | Fixed p95 range |
| --- | ---: | --- | ---: | ---: | ---: | ---: |
| no-private | 1 | rule | 0.015–0.016 | 0.031–0.031 | 0.036–0.036 | 0.055–0.062 |
| no-private | 1 | full | 1.175–1.178 | 1.166–1.169 | 1.742–2.136 | 1.695–1.868 |
| no-private | 1 | full-fresh | 1.110–1.137 | 1.140–1.242 | 1.562–1.597 | 1.690–1.779 |
| no-private | 1000 | rule | 1.323–1.344 | 1.232–1.292 | 1.815–2.656 | 2.059–2.381 |
| no-private | 1000 | full | 36.799–38.120 | 37.268–38.344 | 42.284–42.914 | 45.619–46.772 |
| no-private | 1000 | full-fresh | 39.815–42.613 | 42.472–42.517 | 56.571–68.177 | 56.998–61.821 |
| unrelated | 1 | rule | 0.003–0.003 | 0.011–0.012 | 0.009–0.026 | 0.017–0.019 |
| unrelated | 1 | full | 0.573–0.615 | 0.579–0.584 | 0.674–0.930 | 0.651–0.674 |
| unrelated | 1 | full-fresh | 0.641–0.642 | 0.634–0.647 | 0.704–0.705 | 0.704–0.705 |
| unrelated | 1000 | rule | 0.997–1.023 | 0.886–0.913 | 1.912–2.135 | 1.011–1.965 |
| unrelated | 1000 | full | 33.322–33.628 | 32.981–34.803 | 39.001–41.967 | 37.340–37.859 |
| unrelated | 1000 | full-fresh | 37.746–37.892 | 37.554–37.705 | 40.703–47.497 | 38.367–42.349 |
| positive-first | 1 | rule | 0.004–0.004 | 0.005–0.006 | 0.010–0.013 | 0.017–0.019 |
| positive-first | 1 | full | 0.536–0.555 | 0.556–0.570 | 0.632–1.713 | 1.458–1.679 |
| positive-first | 1 | full-fresh | 0.571–0.594 | 0.586–0.605 | 0.788–0.828 | 0.958–1.051 |
| positive-first | 1000 | rule | 1.249–1.263 | 1.317–1.320 | 1.308–2.338 | 3.479–3.545 |
| positive-first | 1000 | full | 32.827–33.167 | 32.108–32.505 | 40.501–40.872 | 35.557–35.851 |
| positive-first | 1000 | full-fresh | 38.338–39.945 | 37.448–38.572 | 41.352–47.798 | 41.342–53.335 |
| positive-last | 1 | rule | 0.003–0.003 | 0.003–0.003 | 0.003 | 0.004–0.004 |
| positive-last | 1 | full | 0.516–0.548 | 0.502–0.550 | 0.572–0.597 | 0.570–1.720 |
| positive-last | 1 | full-fresh | 0.538–0.586 | 0.536–0.590 | 0.574–0.696 | 0.556–0.632 |
| positive-last | 1000 | rule | 1.036–1.096 | 1.137–1.155 | 1.077–1.984 | 1.990–2.478 |
| positive-last | 1000 | full | 32.097–34.776 | 33.168–34.914 | 38.021–47.832 | 36.891–42.082 |
| positive-last | 1000 | full-fresh | 37.306–38.997 | 37.225–37.291 | 39.916–45.003 | 42.537–43.666 |
| large-cold-negative | 1 | rule | 0.063–0.064 | 0.087–0.088 | 0.095–0.101 | 0.120–0.146 |
| large-cold-negative | 1 | full | 0.895–0.955 | 0.861–0.913 | 1.246–1.484 | 1.271–4.257 |
| large-cold-negative | 1 | full-fresh | 0.978–1.010 | 0.939–0.968 | 1.208–3.484 | 1.190–1.264 |

The unrelated-metadata 1,000-reference rule improves from 0.997–1.023 ms to 0.886–0.913 ms. No universal whole-analyzer speedup is established: complete-analysis timings overlap. The rule’s positive private-member controls are modestly slower, and some single-reference controls pay about 8–24 microseconds for a public-presence lookup/index. The large cold control is 0.063–0.064 ms versus 0.087–0.088 ms in the rule, with overlapping whole-analyzer times. These bounded costs are included above.

## Validation

247 focused tests and type checks passed. After integrating main at 328f46b0, the full suite passed 648 files with 13,186 tests passed and 13 skipped. Complete outputs matched for 165,120 analyzer runs: all 8,256 oracle sources across Excel, Word, PowerPoint and Access with five project contexts (absent, private document, class/public shadowing, private UserForm and ambiguous document names). Cached significant statement tokens and their arrays were frozen in both bundles; no internal errors occurred. Baseline is main at 328f46b0. This does not claim frozen AST or exhaustive repository coverage.
