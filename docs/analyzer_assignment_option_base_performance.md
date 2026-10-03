# Assignment Option Base lookup

Issue: #828. Baseline: `946795a4c0f38e61d8db5497ebb1cf142b91edad`.

Assignment string folding resolved the module option on every array-element read, in addition to shape-provider initialization. Without an option the helper scans the whole module. With conditional activity it first filters all module members, even if the option is first. The assignment rule now lazily resolves the Base once per module/activity invocation, sharing the result across array initialization, reads and procedures. Zero is retained as a cached result. Scalar-only paths do not request it.

No global cache is keyed by parsed nodes: subsequent invocations on the same parsed module with different activity resolve their own option. The existing first-active-option rule and VBA.Array zero-base behavior remain unchanged.

Run `node scripts/benchmark-assignment-option-base.mjs`, then with `--baseline=946795a4c0f38e61d8db5497ebb1cf142b91edad`. Three warmups and fifteen samples on Node 24.18.0 / Ryzen 7 9800X3D. Each sample prepares fresh binding roots and conditional activity outside timing; the script asserts a different root. Parsing is reused. Timings cover the assignment-type rule and its setup, not the whole analyzer. No audit tests/benchmarks ran during timing. Fixtures contain the indicated number of module constants and 1,000 reads of a known string array element. Conditional/explicit fixtures put their Base option before declarations; scalar controls use ordinary numeric assignments.

| Constants/case | Baseline median / p95 ms | Fixed median / p95 ms |
| --- | ---: | ---: |
| 10-default | 4.604 / 5.883 | 4.771 / 5.605 |
| 10-conditional | 10.538 / 13.145 | 9.678 / 12.694 |
| 10-explicit | 3.361 / 5.008 | 3.234 / 4.269 |
| 10-literal-control | 1.646 / 2.563 | 1.685 / 2.375 |
| 100-default | 3.634 / 7.533 | 3.546 / 6.435 |
| 100-conditional | 13.086 / 15.241 | 11.596 / 12.078 |
| 100-explicit | 3.415 / 4.983 | 3.287 / 5.094 |
| 100-literal-control | 1.622 / 2.675 | 1.493 / 2.637 |
| 1000-default | 5.095 / 7.095 | 3.96 / 5.642 |
| 1000-conditional | 51.576 / 59.168 | 42.046 / 53.935 |
| 1000-explicit | 3.985 / 5.562 | 3.976 / 5.796 |
| 1000-literal-control | 1.817 / 3.031 | 1.822 / 3.278 |
| 3000-default | 8.495 / 10.087 | 4.874 / 6.649 |
| 3000-conditional | 167.677 / 201.683 | 139.779 / 186.139 |
| 3000-explicit | 7.063 / 9.485 | 6.421 / 7.453 |
| 3000-literal-control | 2.773 / 4.931 | 2.661 / 5.485 |

Small controls have mixed timing differences: the 10-constant default case rises 4.604 -> 4.771 ms, and the scalar control rises 1.646 -> 1.685 ms. The large conditional fixture still takes 139.779 ms: this removes repeated option resolution, not all conditional/dataflow work. That remaining cost requires a separate audit.

A temporary bundle counter in moduleOptionBase records 1,001 calls before versus 1 after for 1,000 array reads; both runs return the same 1,000 diagnostics. No counter or benchmark-only export is added to production. Tests spy on the public helper to assert exactly one call across two procedures for absent/Base 0/Base 1, and zero calls for scalar-only input.

Validation: type checks; 244 focused tests; full suite (613 files, 12,722 tests passed, 13 skipped); 2,000 complete analyzer outputs matched across six option forms, five array producers, five target types, three indices, module constants, four hosts, activity and multiple procedures. Regression tests also verify exact offending expression spans when active Base changes between invocations on reused parsed nodes, VBA.Array alongside Array, first-active-option semantics and scalar laziness. `git diff --check` passed.
