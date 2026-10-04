# Type references retained with immutable procedures

Type-reference collection cached a complete module but walked every procedure
again when a small edit produced a new module AST. The parser already retains
unchanged procedure nodes with their original source text and absolute spans.
References now follow those nodes in a WeakMap. An edited or shifted procedure
has a new identity and recomputes its references. Cached entries contain only
source facts; project type resolution runs against the current context.

Module aggregation and sorting remain necessary, as does a procedure's first
walk. The cache does not retain source strings, historical module snapshots or
project metadata. Empty results are also reused. Shared reference objects keep
the existing readonly caller contract.

## Reproduction

```powershell
npx vitest run tests/semanticProcedureReferenceWork.test.ts tests/vbaSemanticTokens.test.ts tests/incrementalModuleParse.test.ts
node scripts/benchmark-semantic-procedure-reuse.mjs --baseline=8e85b13f072af5440b68408140bacfd8c2e3ad61
node scripts/benchmark-semantic-procedure-reuse.mjs
```

Set `XLIDE_PERF_WORKBOOK` to the supplied ROneCOne workbook to include the optional
actual-class measurement. It reads the workbook without modifying it or emitting
its source. The fixture is not committed.

The three EOL work-count regressions warm 100 procedures, edit a later body ten
times, require earlier AST identity retention, and compare every reference.
The baseline makes 21,030 statement-token requests; the fix permits only the
edited procedure's 30 requests. A fourth test changes type names, checks shifted
suffix spans, and alternates class/enum resolution on the same source facts.

## Measurement

Node 24.18.0, Ryzen 7 9800X3D. Nine samples after three warmups, separate-process
order baseline/current/current/baseline. Parsing happens before the clock; one
actual collector call on each edited snapshot is timed. Complete results are
checked. The cold control prepends a comment, giving all procedures new identity;
the retained control edits the appended probe body.

| Fixture and scenario | Baseline median ms | Fixed median ms |
| --- | ---: | ---: |
| 1,000 synthetic retained procedures | 3.5114–5.5023 | 1.9519–2.1357 |
| 1,000 synthetic cold procedures | 4.1709–4.4183 | 4.7691–5.0175 |
| ROneCOne, 1,552 retained procedures | 9.9812–13.6311 | 4.6309–4.7185 |
| ROneCOne, cold procedures | 17.7982–20.5503 | 18.7775–19.1088 |

ROneCOne has 969,296 source characters and 5,442 collected references including
the appended probe. The synthetic cold case has a small median cost increase;
actual cold timings are mixed. Fixed cold outliers reached 41.9188 ms for the
synthetic fixture and 49.2938 ms for the actual class. No elimination of cold
costs, tail-latency guarantee or end-to-end editor improvement is claimed.

## Validation

Types and 57 focused tests passed (one existing skipped). Frozen AST/token/trivia
baseline comparison covers 8,256 oracle sources plus 54 early/middle/late edit
snapshots across three EOLs and changing identifier widths. All 24,930 complete
reference/token/lookup records match across project-context transitions;
84,653 semantic tokens were emitted. A retained-prefix identity control also
compares complete references after an appended line. The control was enlarged
from 100 to 200 procedures to exceed the parser's 4,096-character reuse threshold.

Full suite passed: 739 files, 14,575 tests, 33 skipped tests, in 70.41 seconds. This repair contributes to
the remaining large-class latency investigation but does not resolve #985 by
itself. Live editor response and background scheduling still need measurement.
