# Extract Method grouped declaration repair

Issue #886: moving two locals from one Dim statement emitted duplicate whole-line deletion edits. Applying those edits erased the caller's replacement call and following code. Moving only one local deleted its siblings too.

The repair edits each declaration group once. Remaining clauses stay in the caller, and moved clauses keep their original array bounds, As New, fixed string length, type suffix and other declaration syntax. A complete declaration inside the selection is handled within the extracted body; any caller siblings are restored before the call, without adding overlapping source edits or duplicate helper declarations. A selected single output retains its declaration in caller and Function. Deletion spans preserve adjacent colon statements and trailing comments; overlapping separator deletions are unioned.

Validation:

- 39 regression cases failed on baseline 4a6c6ac8; all now pass, across LF, CRLF and CR. They cover full and partial groups, first/middle/last clauses, adjacent colon statements, comments/continuations, array and auto-instantiation syntax, selected groups, input parameters and an output Function.
- 104 focused tests pass, including existing extraction, batching, local ordering and last-statement behavior. Type checking passes.
- 600 frozen-AST ordinary extraction cases retain byte-identical applied text and refusals across the three line endings. Edit lists can coalesce without changing the applied text.
- 600 additional generated grouped/colon/selected-declaration cases preserve expected caller siblings and helper moved declarations, retain the call and following statement, and have non-overlapping source edits.
- Full suite: 643 files passed, 13,084 tests passed, 13 skipped.

These checks exercise source generation and parsed ASTs. They do not execute generated code in the VBE or claim that every Extract Method behavior has been audited.

Performance guard: existing local-order benchmark, Node v24.18.0, AMD Ryzen 7 9800X3D 8-Core Processor           , baseline 4a6c6ac8, 15 rounds after three warmups. Warm reuses the parsed source; fresh adds a trailing source-key comment. Actual Extract Method is timed, result-hash assertions run outside the clock. Fourteen read rows retain complete result hashes. Two write rows change hashes because adjacent deletion edits now coalesce; ordinary applied-text parity is separately verified above. The large read signatures are synthetic analyzer fixtures, not VBE parameter-limit checks. Timings are mixed; this is a correctness repair with no general speedup claim.

| Case | Before median ms | After median ms | Before p95 ms | After p95 ms |
| --- | ---: | ---: | ---: | ---: |
| 1/0/ordered/read/warm | 0.02 | 0.027 | 0.044 | 0.051 |
| 1/0/ordered/read/fresh | 0.073 | 0.078 | 0.391 | 0.303 |
| 5/0/shuffled/read/warm | 0.019 | 0.021 | 0.042 | 0.044 |
| 5/0/shuffled/read/fresh | 0.069 | 0.071 | 0.328 | 0.354 |
| 100/0/shuffled/read/warm | 0.207 | 0.246 | 0.359 | 0.561 |
| 100/0/shuffled/read/fresh | 0.834 | 0.832 | 1.129 | 1.374 |
| 1000/0/ordered/read/warm | 3.68 | 3.706 | 4.862 | 4.659 |
| 1000/0/ordered/read/fresh | 7.415 | 7.451 | 15.229 | 9.909 |
| 1000/8000/shuffled/read/warm | 3.42 | 3.891 | 6.138 | 7.401 |
| 1000/8000/shuffled/read/fresh | 6.176 | 6.171 | 8.534 | 8.983 |
| 1000/100000/shuffled/read/warm | 4.736 | 4.788 | 7.008 | 6.726 |
| 1000/100000/shuffled/read/fresh | 8.313 | 8.503 | 10.352 | 11.087 |
| 3000/100000/shuffled/write/warm | 26.272 | 26.992 | 30.953 | 29.134 |
| 3000/100000/shuffled/write/fresh | 36.459 | 36.29 | 45.582 | 42.317 |
| 1000/100000/shuffled/few/warm | 0.503 | 0.512 | 1.93 | 1.593 |
| 1000/100000/shuffled/few/fresh | 2.791 | 2.876 | 4.681 | 4.307 |

Reproduce sequentially:

```powershell
node scripts/benchmark-extract-method.mjs --local-order --baseline=4a6c6ac8
node scripts/benchmark-extract-method.mjs --local-order
```
