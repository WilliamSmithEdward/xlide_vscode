# Streaming macro string target lookup

`macroNameTarget` previously built every macro completion row, deduplicated the
whole list, converted its Map to an array and filtered that array to locate one
procedure. It now scans the supplied current signatures directly.

A qualified name has one full-name deduplication key: its first non-external
entry determines the answer, so lookup stops there. A bare name retains a Set
of full keys to preserve first-entry deduplication, and returns no target once
two distinct candidate keys match. It no longer allocates completion row objects
or a filtered candidate array. Completion enumeration itself is unchanged.

The lookup keeps original procedure identity, external exclusions, case folding,
quoted workbook prefixes, empty macro override precedence and full-key collision
behavior even for dotted signature names. Caller-owned arrays are read fresh;
no index or cache was introduced.

## Reproduction

```powershell
npx vitest run tests/macroTargetLookupWork.test.ts tests/macroNames217.test.ts
node scripts/benchmark-macro-target-lookup.mjs --baseline=1b445e849ef18382680c4d7bb71a49de7814447e
node scripts/benchmark-macro-target-lookup.mjs
```

Three work-count cases place the qualified target first, middle and last in
1,000 signatures. All fail on the baseline because every procedure is read;
they pass with the bound of the target position plus two name reads. Four
output controls pass on both, covering duplicates, exclusions, ambiguity,
full-key collisions and caller-array mutation.

## Measurement

Node 24.18.0, Ryzen 7 9800X3D; nine samples after three warmups, separate-process
order baseline/current/current/baseline. The table is milliseconds for ten
actual resolver calls over 10,000 signatures. Fixture creation and result
identity assertions are outside the clock. Parsing and editor work are absent.

| Target | Baseline median ms | Fixed median ms |
| --- | ---: | ---: |
| Qualified first | 8.1642–8.1994 | 0.0014 |
| Qualified middle | 8.7310–9.0399 | 0.7696–0.9821 |
| Qualified last | 8.6094–8.7469 | 1.7396–2.0614 |
| Qualified missing | 8.5251–8.8177 | 1.5663–1.5882 |
| Bare unique | 8.6062–9.2063 | 7.2923–7.3595 |
| Bare ambiguous, duplicate at end | 8.7793–8.9524 | 7.2078–7.4536 |

Small cases remain sub-millisecond batch work. Bare lookup still scans and
allocates deduplication keys; qualified late/missing lookup still scans all
relevant signatures. This does not establish an end-to-end hover improvement
or removal of cold editor costs.

## Validation

Types and 26 focused tests pass. A deterministic baseline/current differential
uses 3,000 frozen signature contexts and 135,000 queries, including empty and
dotted names, Unicode, case folding, external records, duplicate full keys,
quoted prefixes and override/fallback arrays. Every returned object identity
matches (15,421 resolved targets), and complete completion candidate arrays
match. Separate mutable caller controls remain fresh. No shared input mutation
occurs. Full suite passed: 740 files, 14,582 tests, 33 skipped tests, in 75.15 seconds.
