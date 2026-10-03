# Object-value read token scan performance

valueReads in objectValues.ts previously scanned every accumulated read with out.some for each candidate token. This is redundant: whole Let values and isolated Print items have no adjacent scalar operator; the operator loop visits each lexer token once. Tokens are distinct objects allocated by the lexer. Operator reads therefore cannot duplicate either those initial whole-value reads or an earlier operator read. Removing the check preserves whole-value reads before operator reads in output order, without adding a Set.

The late-bound Collection path also called noDefaultReads and then used toks.indexOf for each result merely to discard indexed reads. The private helper now accepts includeIndexed (true by default). This caller requests false, skipping indexed reads and their parenthesis matching directly. Other callers retain indexed-type diagnostics. Whole builtin arguments and condition reads keep their existing order.

## Reproduce

```powershell
node scripts/benchmark-object-read-token-scans.mjs --baseline=1ff0c6be22e47e0a112ed8dbb58842464b813e8c
node scripts/benchmark-object-read-token-scans.mjs
```

The baseline option substitutes only objectValues.ts when bundling; the surrounding code is identical. Parse once, build fresh symbols before each sample outside the timer, asserting distinct root identity. Time rule setup and all procedure statement/header visits. Three warmups and 15 samples; the script reports median and p95. Run without other benchmarks/tests in parallel.

Fixtures declare n As Long and x As Object (unset), then use Debug.Print with the requested number of n operands joined by +, separate n items joined by semicolons, or CStr(x) arguments joined by +. Groups of 75 items are joined with line continuations. At 1,500 reads there are 19 continuations and physical lines stay below 1,023 characters; the harness asserts every requested name remains in the parsed Print statement. These source constraints do not constitute an external VBA runtime/compiler oracle check. Every fixture must produce zero findings from this rule. scalar-control is always the same one-assignment procedure; its repeated rows expose timing variation rather than scaling.

## Results

2026-10-03, Node v24.18.0, AMD Ryzen 7 9800X3D 8-Core Processor. Rule medians in milliseconds:

| Reads and mode | Before | After |
| --- | ---: | ---: |
| 1-operators | 0.023 | 0.017 |
| 1-print-items | 0.015 | 0.010 |
| 1-value-arguments | 0.030 | 0.019 |
| 1-scalar-control | 0.012 | 0.007 |
| 10-operators | 0.015 | 0.014 |
| 10-print-items | 0.016 | 0.014 |
| 10-value-arguments | 0.038 | 0.038 |
| 10-scalar-control | 0.009 | 0.008 |
| 100-operators | 0.061 | 0.047 |
| 100-print-items | 0.060 | 0.037 |
| 100-value-arguments | 0.163 | 0.112 |
| 100-scalar-control | 0.005 | 0.006 |
| 500-operators | 0.320 | 0.209 |
| 500-print-items | 0.284 | 0.202 |
| 500-value-arguments | 0.462 | 0.434 |
| 500-scalar-control | 0.004 | 0.007 |
| 1500-operators | 1.415 | 0.356 |
| 1500-print-items | 1.448 | 0.520 |
| 1500-value-arguments | 2.601 | 0.979 |
| 1500-scalar-control | 0.004 | 0.004 |

The 1,500-operand baseline performs 1,124,250 redundant token-identity comparisons; the changed helper performs zero. The 1,500-read cases improve from 1.415 to 0.356 ms (operators), 1.448 to 0.520 ms (Print items), and 2.601 to 0.979 ms (value arguments). Small scalar controls vary in both directions (500-read row: 0.004 to 0.007 ms); these are rule-only synthetic timings, not end-to-end or typical-workbook latency claims. Other helper and diagnostic work remains, so the whole rule is not claimed to have constant cost.

## Validation

- Four new tests preserve single whole Let reads, whole Print items before operator reads, separate repeated operand spans, and late-bound indexed exclusions versus whole builtin and typed indexed reads.
- Eight focused files: 66 tests passed; typecheck passed.
- Full suite: 605 files, 12,629 tests passed, 13 skipped.
- 1,500 complete analyzeModule outputs match baseline 1ff0c6be exactly across five hosts, ten types, ten forms and three scope configurations with alternating VBA7 activity; neither build reported an internal error.
- 5,000 generated mixed/malformed token streams yield 15,000 exact private-helper comparisons (valueReads, default noDefaultReads, and whole-object-only noDefaultReads against the old filtered result). Names, comments, Unicode, operators, members, labels, parentheses and conditional text are included. Helpers are exposed only in the scratch validation bundle; no public API was added.
