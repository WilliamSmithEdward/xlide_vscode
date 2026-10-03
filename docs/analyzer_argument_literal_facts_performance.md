# Lazy literal facts for argument validation

checkArgumentTypes previously requested knownLocalLiteralValuesAt during every procedure visitor's setup. The fact provider memoizes by binding context, but a first request constructs whole/local/reaching value state. No-call and literal-only procedures do not use these callbacks, and warm bindings hide this setup cost.

Keep the existing valuesAt callback but initialize its provider only on the first actual query, retaining the provider within that visitor. Named values, Variant scalar-to-object checks and Date serial fallbacks still receive their preceding and active-branch facts. No global cache or invalidation policy changes.

## Reproduction

```powershell
node scripts/benchmark-argument-literal-facts.mjs --baseline=5b1120c5 --rounds=15
node scripts/benchmark-argument-literal-facts.mjs --rounds=15
```

The benchmark parses once, then creates a new bound symbol root before every warmup/sample. It asserts those root identities differ. Binding is outside the timed region; the timed work includes the public argument-type factory and procedure statement visitors. Fixtures declare 25 Long locals plus a Variant holding 300, then perform repeated assignments. The no-call case adds no call, the named control passes the Variant to a Long parameter, and the literal case passes a nonnumeric string and requires one mismatch. The named control genuinely needs literal facts.

Node 24.18.0, AMD Ryzen 7 9800X3D; three warmups and 15 samples, median milliseconds:

| Assignments | No calls before | After | Named argument before | After | Literal argument before | After |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 100 | 0.345 | 0.188 | 0.258 | 0.498 | 0.180 | 0.086 |
| 1,000 | 1.167 | 0.596 | 0.970 | 1.173 | 1.313 | 0.485 |
| 3,000 | 2.519 | 1.421 | 2.690 | 2.643 | 2.549 | 1.341 |

The named control still computes the facts. Its smaller cases measured slower; this change does not promise faster needed-fact validation. Absolute values vary with load and execution order. An earlier run measured the largest no-call case at 2.607/1.623 ms and the largest named control at 2.751/2.693 ms. Final runs followed validation and included the fresh-root assertion.

These are isolated cold-rule stress measurements, not editor latency. Other analyzer rules may already populate or require the shared facts, so whole-analyzer savings can be smaller or absent. Eager held-object preparation remains in this baseline; the independent held-object deferral PR #806 is compatible and addresses a different provider.

## Validation

- Type checking passed; full suite: 602 files, 12,556 tests passed, 13 skipped.
- 82 targeted tests passed.
- 1,500 complete analyzer outputs matched baseline 5b1120c5 across numeric/string/Date/object/Variant parameters, known and unknown values, call forms, branches and conditional compilation.
- Four regressions require zero literal-provider requests for unused callbacks, one for repeated named Variant values, preserved scalar-to-object mismatches and independent activity environments on reused parsed nodes.
