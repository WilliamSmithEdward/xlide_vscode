# Set form/control metadata lookup

Issue: #824. Baseline: `946795a4c0f38e61d8db5497ebb1cf142b91edad`.

The Set target and value paths repeatedly searched the same project form surface and then the same control metadata. The rule now resolves the current form lazily once per invocation and caches only queried control names, including missing results. The cache is shared across procedures within that invocation, and discarded afterward. It is not a global cache keyed by mutable project metadata.

The first matching userform remains authoritative, even if incomplete or missing the queried control. A same-name surface of another kind is skipped. Within that form, the first same-name member whose return type begins with MSForms. remains authoritative; a same-name non-control is skipped. Source declaration guards and local value shadowing remain in place. The target and value paths share the query cache.

Run `node scripts/benchmark-set-form-controls.mjs`, then with `--baseline=946795a4c0f38e61d8db5497ebb1cf142b91edad`. Three warmups and fifteen samples on Node 24.18.0, Ryzen 7 9800X3D. Each sample prepares fresh symbol bindings outside timing and asserts a new root. Timing includes rule/procedure setup and a whole statement walk; parsing is outside timing. No audit tests/benchmarks ran concurrently during measurement. Fixtures use 1,000 repeated Set statements. Form-class-value places the current form after the indicated unrelated form surfaces; other form modes place the queried member after that many unrelated controls. The scalar control has no form context. These are isolated Set-rule stress timings, not complete-analyzer speedup measurements.

| Size/case | Baseline median / p95 ms | Fixed median / p95 ms |
| --- | ---: | ---: |
| 10-form-class-value | 1.111 / 1.654 | 0.884 / 1.259 |
| 10-form-member-value | 0.843 / 0.968 | 0.638 / 0.697 |
| 10-form-member-target | 0.59 / 0.943 | 0.453 / 0.932 |
| 10-form-missing-value | 1.611 / 1.89 | 0.983 / 1.786 |
| 10-scalar-control | 0.401 / 0.987 | 0.418 / 0.652 |
| 100-form-class-value | 1.318 / 1.797 | 0.594 / 0.869 |
| 100-form-member-value | 1.527 / 1.619 | 0.539 / 0.582 |
| 100-form-member-target | 0.625 / 0.884 | 0.228 / 0.404 |
| 100-form-missing-value | 3.747 / 4.129 | 0.838 / 1.226 |
| 100-scalar-control | 0.367 / 0.553 | 0.354 / 0.675 |
| 1000-form-class-value | 7.683 / 8.272 | 0.531 / 0.814 |
| 1000-form-member-value | 11.458 / 12.07 | 0.543 / 0.661 |
| 1000-form-member-target | 4.53 / 5.001 | 0.234 / 0.382 |
| 1000-form-missing-value | 31.252 / 33.823 | 0.853 / 0.911 |
| 1000-scalar-control | 0.345 / 0.659 | 0.351 / 0.63 |
| 3000-form-class-value | 24.892 / 25.961 | 0.556 / 0.946 |
| 3000-form-member-value | 32.806 / 33.729 | 0.547 / 0.671 |
| 3000-form-member-target | 13.997 / 15.239 | 0.246 / 0.469 |
| 3000-form-missing-value | 94.098 / 97.275 | 0.841 / 1.45 |
| 3000-scalar-control | 0.37 / 0.588 | 0.353 / 0.785 |

Scalar controls have mixed small timing differences (10 case: 0.401 -> 0.418 ms; 3,000 case: 0.370 -> 0.353 ms). This change removes repeated metadata searches; it does not eliminate other scans or inference costs for unrelated paths. The lazy query cache avoids eagerly indexing unused controls.

Actual getter counts over 100 unrelated forms and 100 unrelated members, with 200 target/value assignments, fall from 20,000 form-name reads and 20,000 member-name reads to 100 of each. All runs return the same 200 diagnostics. Regression tests assert generous deterministic cost bounds instead of timing thresholds.

Validation: type checks; 244 focused tests; 2,500 complete analyzer outputs matched the baseline across four hosts, standard/form module contexts, local/module/enum/const/function shadows, arrays, conditional activity, mixed member types and duplicate/empty/incomplete form surfaces. Seven targeted tests cover metadata read counts, misses, first-match behavior, case folding, metadata mutations between invocations on the same array/object, source shadowing and unused metadata laziness.

The final full suite passed: 613 files, 12,722 tests passed and 13 skipped. The final baseline already includes both neighboring fixes, so the table measures only this form/control cache change. `git diff --check` passed.
