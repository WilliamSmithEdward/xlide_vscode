# Set target declaration membership

Issue: #822. Baseline: `041fa29f6c61bd5b55bb5290a6de962b0f4a3405`.

The Set visitor copied every procedure/module child into a combined array and scanned declaration names for every Set statement. The check only affects form-control targets, but it also ran with no form context. It now skips that work without a form context. In forms, exact child-name sets are built lazily: one procedure set per visitor and one shared module set per rule invocation. They are not retained across invocations.

This preserves the original direct-declaration semantics, including case folding and procedure-local isolation. The broader runtime shadow index is not interchangeable: it also includes enum members, return bindings and project-visible names, which this control check previously excluded. No form/control metadata lookup behavior changed.

Run `node scripts/benchmark-set-source-names.mjs`, and run it with `--baseline=041fa29f6c61bd5b55bb5290a6de962b0f4a3405` for comparison. Measurements cover the entire Set visitor including its invocation/procedure setup and statement walk. Parsing and fresh symbol binding are outside timing; the script asserts a different binding root for each sample. Three warmups, fifteen samples on Node 24.18.0, Ryzen 7 9800X3D; no other audit tests/benchmarks ran during these measurements. Each fixture has 1,000 Set statements and the indicated number of unrelated locals. Early-local places the target first; late-local places it last. Form-local has a form context and a declared target. The no-Set control uses ordinary assignments. These isolated rule timings do not represent whole-analyzer speedups.

| Locals/case | Baseline median / p95 ms | Fixed median / p95 ms |
| --- | ---: | ---: |
| 10-late-local | 0.908 / 1.327 | 0.754 / 1.02 |
| 10-early-local | 0.563 / 1.146 | 0.587 / 0.78 |
| 10-form-local | 0.459 / 0.505 | 0.456 / 0.612 |
| 10-no-set-control | 0.122 / 0.147 | 0.122 / 0.172 |
| 100-late-local | 1.288 / 2.015 | 0.498 / 1.059 |
| 100-early-local | 0.5 / 1.076 | 0.376 / 0.821 |
| 100-form-local | 1.103 / 1.775 | 0.611 / 0.751 |
| 100-no-set-control | 0.084 / 0.366 | 0.115 / 0.133 |
| 1000-late-local | 6.966 / 7.583 | 0.481 / 1.323 |
| 1000-early-local | 1.525 / 3.989 | 0.486 / 1.272 |
| 1000-form-local | 7.331 / 7.949 | 0.519 / 2.205 |
| 1000-no-set-control | 0.163 / 0.201 | 0.227 / 0.298 |
| 3000-late-local | 21.564 / 23.013 | 0.754 / 1.093 |
| 3000-early-local | 3.086 / 3.542 | 0.743 / 0.837 |
| 3000-form-local | 21.283 / 22.531 | 0.827 / 0.86 |
| 3000-no-set-control | 0.324 / 1.016 | 0.328 / 0.346 |

Small controls have mixed timing differences, including the no-Set control (1,000 locals: 0.163 -> 0.227 ms) and the 10-local early target (0.563 -> 0.587 ms). The confirmed improvement removes repeated declaration work on Set statements, not every setup cost in this rule.

A getter-based count over actual symbol objects, 100 unrelated locals and 100 Set statements, recorded 10,508 name reads before the change versus 408 without a form context and 509 within one. All runs returned the same 100 diagnostics. Regression tests assert a generous bound on actual reads, avoiding timing thresholds.

Validation: type checks; 243 focused tests; full suite (610 files, 12,703 passed, 13 skipped); 2,000 complete analyzer outputs matched the baseline across four hosts, form/standard module contexts, local/module/enum/const/function shadows, mixed control metadata, arrays and conditional activity. Regression tests explicitly preserve project-visible and enum-member exclusions, procedure-local isolation, conditional declaration refresh, scalar/array Set behavior and case-insensitive shadowing.
