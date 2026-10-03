# Assignment object-type facts performance

checkAssignmentTypes repeatedly asked whether the same declared type was an object, what its default accepts, whether that default holds an object, and whether it is read-only. Project recognition filters class surfaces, and verdict/read-only checks scan surfaces and default members. A sequence of assignments to one declared type repeats those scans.

Keep one record per exact declared-type spelling during this rule invocation. Record non-object/unresolved types too. Positive records retain the existing verdict, holding-default and read-only-default results, in the existing precedence (holding defaults win before read-only defaults). Scalar and array-element callers continue their distinct diagnostic decisions. State/value checks are not cached here: objectLetStateAt still runs at each applicable statement, so an unset object and a later set object report different errors.

Caches start fresh on every invocation and are shared across its procedures. Changed project metadata, host context, conditional bindings and declaration types are read on the next invocation. Keys preserve the original raw type spelling; message labels are still built from each assignment's declared type. Extra retained space is O(distinct declared types queried) for this invocation. The first array-element query may compute static default facts that only a scalar query would consume; they are not used to add diagnostics to the element path. No global cache or public API was added, and first-match/ambiguity semantics inside existing helpers remain unchanged.

## Reproduce

```powershell
node scripts/benchmark-assignment-object-facts.mjs --baseline=bc33fc75cfe836fdb7ec5be381526232c69c9e13
node scripts/benchmark-assignment-object-facts.mjs
```

The baseline option substitutes only assignments.ts in the bundle; surrounding code is identical. Parse once, build fresh symbols before each sample outside timing, assert distinct root identity, then time checkAssignmentTypes. Three warmups and 15 measured samples; the script reports median and p95. Run without other tests/benchmarks in parallel.

Each fixture supplies the stated number of preceding project classes then the target class (absent in missing). It repeats 1,000 assignments to a required-default class, a read-only zero-argument default, a writable zero-argument default, an unresolved class, or an element of a required-default class array. Required/read-only/element cases must produce 1,000 findings; writable/missing cases none. The scalar control performs 1,000 n = 1 assignments with n As Long and the same metadata pool. Binding/parsing and fixture construction are excluded.

## Results

2026-10-03, Node v24.18.0, AMD Ryzen 7 9800X3D 8-Core Processor. Rule medians in milliseconds:

| Preceding classes and mode | Before | After |
| --- | ---: | ---: |
| 10-required | 1.564 | 0.817 |
| 10-readonly | 1.260 | 0.409 |
| 10-writable | 1.253 | 0.371 |
| 10-missing | 1.897 | 1.535 |
| 10-element-required | 1.318 | 0.751 |
| 10-scalar-control | 1.731 | 1.655 |
| 100-required | 2.279 | 0.451 |
| 100-readonly | 2.688 | 0.338 |
| 100-writable | 2.737 | 0.313 |
| 100-missing | 2.167 | 1.501 |
| 100-element-required | 2.311 | 0.690 |
| 100-scalar-control | 1.769 | 1.712 |
| 1000-required | 14.726 | 0.475 |
| 1000-readonly | 18.528 | 0.363 |
| 1000-writable | 18.351 | 0.341 |
| 1000-missing | 6.215 | 1.500 |
| 1000-element-required | 14.269 | 0.619 |
| 1000-scalar-control | 1.714 | 1.672 |
| 3000-required | 44.814 | 0.507 |
| 3000-readonly | 87.408 | 0.424 |
| 3000-writable | 95.240 | 0.399 |
| 3000-missing | 23.039 | 1.512 |
| 3000-element-required | 66.384 | 0.691 |
| 3000-scalar-control | 2.880 | 1.677 |

These are synthetic rule-only cases, not typical project sizes or end-to-end analyzer latency claims. At 3,000 preceding classes, required defaults improve 44.814 to 0.507 ms, read-only 87.408 to 0.424 ms, writable 95.240 to 0.399 ms, missing classes 23.039 to 1.512 ms and required-default array elements 66.384 to 0.691 ms. Scalar controls also varied (10 classes: 1.731 to 1.655 ms; 3,000: 2.880 to 1.677 ms). First type queries and other expression/binding/state work remain; this does not optimize every type-inference caller. The earlier 3,000 read-only probe was about 93 ms before versus 87 ms in the final run.

Stable getter counters on 100 repeated assignments and 100 preceding classes independently verify work:

| Metadata reads | Before | After |
| --- | ---: | ---: |
| Read-only: class names | 40,000 | 400 |
| Read-only: default flags | 20,000 | 200 |
| Missing type: class names | 10,000 | 100 |
| Array element: class names | 30,000 | 300 |
| Array element: default flags | 10,000 | 100 |

## Validation

- Six new tests cover actual metadata read budgets, same-array/object updates on reused parsed nodes, scalar/element distinction for one type, state-dependent 91 versus 438, conditional local shadowing, and Word defaults holding objects versus read-only defaults.
- Six focused files: 250 tests passed; typecheck passed.
- Full suite: 607 files, 12,638 tests passed, 13 skipped.
- 2,000 complete analyzeModule outputs exactly match baseline bc33fc75 with no internal errors: required/optional/ParamArray/zero-argument defaults, writable/setter flags, duplicate/incomplete/mixed-kind surfaces, qualified and array declarations, local shadowing, changing object states, ten assignment forms, four host contexts and conditional activity. Messages, spans and order match.
