# Analyzer symbol association performance — 2026-10-03

buildModuleSymbols repeated a full root-symbol search per procedure when collecting DefType implicit locals, and a full name search per module-level member attribute. The fix builds indexes within each symbol construction call. Offset lookup preserves the first match, while name lookup preserves all same-name accessors/duplicates and attribute order. No symbol index survives the call, and modules without member attributes skip the attribute index.

Run `node scripts/benchmark-symbol-association.mjs --rounds=15` in each checkout. Measurements use the same parsed module per sample (parsing excluded), three warmups, 15 samples, Node 24.18.0 and AMD Ryzen 7 9800X3D. Baseline: main after PR #692. Median milliseconds:

| Workload | Before | After |
| --- | ---: | ---: |
| DefType procedures: 500 | 2.006 | 0.816 |
| DefType procedures: 2,000 | 4.746 | 2.469 |
| DefType procedures: 5,000 | 24.370 | 6.350 |
| Attributed procedures: 500 | 1.825 | 0.203 |
| Attributed procedures: 2,000 | 23.657 | 0.677 |
| Attributed procedures: 5,000 | 221.631 | 1.622 |

Synthetic stress cases show scaling and isolate the repeated association work; they are not end-to-end editor latency predictions. Attribute arrays retain their existing concatenation behavior.

Regression tests count procedure name-span reads and exercise same-name property accessors, case-insensitive member attributes, missing targets, source attribute order, module attributes, and Option Explicit/DefType guards. Type checking, focused tests and extension compilation passed; full-suite results are recorded in the PR.
