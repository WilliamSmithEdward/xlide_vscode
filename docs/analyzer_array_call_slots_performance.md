# Fixed-array call slot eligibility

The fixed-array subscript scanner previously parsed argument slots for every unqualified named call whenever any array shape was known. For an ordinary call with no applicable declaration and no loop-counter facts, both reporting paths were ineligible. Nested calls still repeatedly sliced and scanned their argument suffixes.

Skip argument-slot parsing when neither a declared shape nor a nonempty counter map is available. Continue scanning the original token stream so a real array access inside an unrelated call is still checked. Keep the existing symbolic counter path whenever counter facts are present, including array parameters whose declared bounds are unknown.

## Reproduction

```powershell
node scripts/benchmark-array-call-slots.mjs --baseline=5fa30a3c --rounds=15
node scripts/benchmark-array-call-slots.mjs --rounds=15
```

The benchmark times the public fixed-array rule with parsing and binding outside the timed region, three warmups, 15 samples and warmed token caches. The unrelated fixture declares one fixed array but nests ordinary F calls around a literal. The inner-array fixture instead nests them around a(5), requiring one out-of-bounds diagnostic. The control has no fixed array and requires no diagnostic.

Node 24.18.0, AMD Ryzen 7 9800X3D; median milliseconds:

| Nested calls | Unrelated before | After | Inner array before | After | No-array control before | After |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 10 | 0.060 | 0.032 | 0.034 | 0.020 | 0.006 | 0.006 |
| 100 | 0.478 | 0.039 | 0.474 | 0.031 | 0.009 | 0.010 |
| 1,000 | 16.044 | 0.149 | 20.067 | 0.125 | 0.036 | 0.031 |

These isolated-rule stress fixtures do not measure editor latency or establish valid VBA nesting limits. Oversized inputs still matter while editing or analyzing generated text. Absolute timings vary with machine load.

## Validation

- 85 targeted tests passed.
- 4,000 generated fixtures matched complete diagnostic callback outputs against baseline 5fa30a3c: nested ordinary calls, declared and value-built arrays, symbolic/unknown loop bounds, malformed arguments, wrong dimensions, member calls, intrinsic calls and conditional activity.
- Token rawText reads for 200 nested ordinary calls fell from 168,243 to 7,644. The regression test bounds those reads independently of wall-clock timing.
- Four regressions verify eligibility cost, genuine inner-array violations, symbolic bounds of array parameters and wrong-dimension reports inside ordinary calls.

- Type checking passed; full suite: 597 files, 12,532 tests passed, 13 skipped.
