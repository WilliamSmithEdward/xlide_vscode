# Statement-form class Sub candidate filtering

Qualified assignment values previously resolved every receiver just to ask whether its member was a Sub of a project class. projectClassMemberAt can return that member only from kind=class project metadata. An incremental, lazy set of lowercase Sub names per checkStatementForms invocation now rejects impossible names before resolving the receiver. Collection stops on a matching candidate and resumes where it left off on the next unknown name; known candidates reuse the set. Candidate names still use the original resolver, so duplicate members, ambiguous receivers and access/type behavior keep its semantics. No persistent metadata cache is added.

The actual-rule tests show 1,000 resolver calls reduced to zero with absent metadata, property-only metadata, unrelated Sub names and standard-module-only metadata. Real class Subs still report; a Function preceding a duplicate Sub still wins; a new rule invocation observes changed member metadata. The final nine tests have five baseline failures and four passing controls. They also verify that an early hit leaves unrelated Sub metadata unread and that traversal resumes correctly after hits and misses.

## Reproduce

Baseline 1f8a0bf1d1d54c8540a8161176f327ece228c7bf; Node v24.18.0. Reuse the existing pinned benchmark implementation:

```powershell
node scripts/benchmark-analyzer-statement-labels.mjs --baseline=1f8a0bf1d1d54c8540a8161176f327ece228c7bf --rule-only > rule-before.json
node scripts/benchmark-analyzer-statement-labels.mjs --rule-only > rule-after.json
node scripts/benchmark-analyzer-statement-labels.mjs --baseline=1f8a0bf1d1d54c8540a8161176f327ece228c7bf > before.json
node scripts/benchmark-analyzer-statement-labels.mjs > after.json
```

Each run uses 15 measured rounds after three warm-ups. Parsing/tokenization and rule symbol construction occur outside the clock. Fresh mode uses a unique trailing comment and fresh parsed procedure bodies. Assertions run outside the clock. Before/after runs were sequential, with no own concurrent tests/probes. Other Python jobs were running on the machine, so these timing samples are contended. Controls and whole-analyzer results are mixed, including slower arithmetic workloads: no universal improvement or regression attribution is established. Deterministic skipped resolver calls are the primary evidence. The qualified rule workload uses 1,000 Foo.Bar assignments and standard-module metadata with no class Sub names.

The additional controls use 1,000 classes with 100 members each (100,000 total), testing one and 1,000 missing or real Sub references. Reproduce with the same script plus --rule-only --sub-member-controls. A single negative reference now spends 0.198 ms rather than 0.026 ms scanning metadata; 1,000 negatives improve from 17.689 to 1.110 ms. This upfront scan is an explicit tradeoff. An early positive no longer collects unrelated metadata; positive timings are mixed. The cursors and seen-name set are discarded after the rule invocation.

## Rule only

Milliseconds per invocation.

| Workload | Median before | Median after | p95 before | p95 after |
| --- | ---: | ---: | ---: | ---: |
| assignments-warm | 0.339 | 0.555 | 1.016 | 3.673 |
| assignments-fresh | 0.528 | 0.633 | 1.317 | 1.218 |
| long-body-warm | 1.283 | 1.931 | 1.918 | 2.775 |
| long-body-fresh | 2.352 | 3.139 | 4.524 | 5.175 |
| project-no-hit-warm | 0.312 | 0.351 | 0.412 | 0.367 |
| project-no-hit-fresh | 0.347 | 0.585 | 2.014 | 0.767 |
| qualified-warm | 16.904 | 0.904 | 25.247 | 1.285 |
| qualified-fresh | 19.076 | 1.042 | 25.948 | 1.643 |
| label-hit-warm | 0.864 | 1.115 | 1.177 | 1.348 |
| label-hit-fresh | 0.906 | 1.131 | 3.697 | 3.698 |
| short-warm | 0.003 | 0.006 | 0.009 | 0.019 |
| short-fresh | 0.004 | 0.008 | 0.009 | 0.026 |

## Full analyzer

Milliseconds per invocation.

| Workload | Median before | Median after | p95 before | p95 after |
| --- | ---: | ---: | ---: | ---: |
| many-warm | 60.573 | 75.100 | 183.803 | 96.578 |
| many-fresh | 80.892 | 84.748 | 109.660 | 122.892 |
| body-warm | 226.740 | 127.647 | 281.224 | 142.511 |
| body-fresh | 292.443 | 158.696 | 357.417 | 499.813 |
| branches-warm | 118.392 | 57.537 | 227.480 | 106.305 |
| branches-fresh | 89.503 | 54.466 | 144.940 | 63.379 |
| types-warm | 43.389 | 25.480 | 67.495 | 55.534 |
| types-fresh | 60.618 | 27.624 | 184.491 | 35.424 |
| loops-warm | 86.847 | 46.224 | 127.181 | 50.401 |
| loops-fresh | 78.097 | 42.810 | 360.499 | 49.861 |
| short-warm | 0.617 | 0.398 | 31.133 | 0.489 |
| short-fresh | 0.487 | 0.335 | 0.682 | 0.442 |
| qualified-warm | 16.998 | 9.185 | 63.340 | 11.838 |
| qualified-fresh | 18.156 | 12.842 | 56.681 | 46.375 |
| label-hit-warm | 32.329 | 23.942 | 58.647 | 28.012 |
| label-hit-fresh | 40.526 | 23.608 | 91.007 | 28.530 |

## Large metadata controls

Milliseconds per invocation.

| Workload | Median before | Median after | p95 before | p95 after |
| --- | ---: | ---: | ---: | ---: |
| large-project-negative-short-warm | 0.026 | 0.198 | 0.048 | 0.229 |
| large-project-negative-short-fresh | 0.020 | 0.185 | 0.062 | 0.197 |
| large-project-negative-many-warm | 17.689 | 1.110 | 28.521 | 1.927 |
| large-project-negative-many-fresh | 18.766 | 1.548 | 51.198 | 3.267 |
| large-project-sub-short-warm | 0.009 | 0.023 | 0.016 | 0.045 |
| large-project-sub-short-fresh | 0.009 | 0.023 | 0.026 | 0.034 |
| large-project-sub-many-warm | 17.230 | 20.495 | 21.541 | 26.615 |
| large-project-sub-many-fresh | 20.064 | 19.482 | 27.603 | 21.772 |

Final native validation: 8,256 string-source oracle cases × four hosts (Excel, Word, PowerPoint, Access) × four project metadata contexts = 132,096 complete diagnostic comparisons. Default, standard-module-only, class Sub/property, and duplicate Function-before-Sub metadata contexts all match baseline exactly, including diagnostic order/data. Statement-token arrays and each token are frozen in both bundles. Zero internal errors. Non-string corpus entries are excluded. The final nine tests and 92 focused tests pass, along with type checks; the full suite on the branch integrating #878 passes: 638 files, 12,981 tests passed, 13 skipped.
