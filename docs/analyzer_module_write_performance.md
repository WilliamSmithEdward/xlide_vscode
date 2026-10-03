# Module-write enclosing-callee lookup

The shared write scanner treats whole names passed to unknown calls as possible ByRef writes. Previously every potential argument searched backward through its token prefix to find the enclosing callee. Wide parenthesized calls and bare argument lists repeated that work.

The scanner now advances a cursor through each statement segment only as far as needed, maintaining a stack of callee names for unmatched opening parentheses. Queries come in increasing token order. Nested calls, runtime functions, source shadowing, bare calls and unmatched closing parentheses retain their previous results. The stack is local to the segment and does not persist across analyses.

## Measurement

Run `node scripts/benchmark-module-write-callees.mjs --rounds=15`; compare with `--baseline=217f53c3`. The baseline loader substitutes only the old module-state implementation. Lexing happens outside timing via the existing token cache. Each fixture supplies N variable arguments to an unknown call, a runtime call, or a bare call and verifies the write count.

Node 24.18.0, AMD Ryzen 7 9800X3D, three warmups / 15 samples. Values are median milliseconds:

| Arguments | Unknown before | Unknown after | Runtime before | Runtime after | Bare before | Bare after |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| 100 | 0.061 | 0.038 | 0.046 | 0.025 | 0.051 | 0.035 |
| 1,000 | 2.324 | 0.147 | 1.944 | 0.086 | 1.959 | 0.209 |
| 10,000 | 205.821 | 1.288 | 329.041 | 0.904 | 320.455 | 1.362 |

The larger fixtures are editor-input stress cases, not claims about valid VBA call sizes. These measurements cover the isolated write scanner, not end-to-end editor latency. The operation-count fixture with 1,000 arguments drops from 2,013,005 raw-token text reads to 14,006; the regression budget fails the old implementation and passes this change.

## Validation

10,000 generated statement streams produce exactly the same written-name sets in the same insertion order as the old scanner, including malformed parentheses, statement separators, assignment prefixes, nested unknown/runtime calls and source declarations. Targeted tests cover the token-read budget, nested callees, receiver/bare calls, parenthesized arguments, runtime shadowing and incomplete parentheses.

Validation completed: 39 targeted tests and the type check passed; full suite 583 files / 12,282 tests passed (13 skipped).
