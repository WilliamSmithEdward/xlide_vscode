# Small-edit lexer reuse

Related: https://github.com/WilliamSmithEdward/xlide_vscode/issues/985

Typing in a large module previously invalidated the lexer cache and rebuilt every
token, even when only one identifier character changed. Completion, hover and
semantic coloring all consume the cached lexer.

The cache now tries the newest source of a compatible size. Small expression
lookups do not hide the active module. A small edit without physical newline
changes re-lexes the surrounding logical line, verifies its terminating newline,
and reuses the other tokens. Offset changes clone suffix tokens and trivia rather
than mutating previous snapshots.

Sources below 4 KiB, replacements over 128 characters, physical newline changes,
split CRLF boundaries, logical windows over 16 KiB, and a lost terminating newline
use the full lexer. Cold sources still use the full lexer. The existing eight-entry
cache stays bounded. No response delay was added.

## Reproduce

Run the portable correctness checks:

```powershell
npx vitest run tests/incrementalLexerPerformance.test.ts
```

The optional benchmark reads ROneCOne from the supplied workbook without modifying
it. It compares cached incremental lexing with a full lex of the same edited source
and checks complete token equality at early, middle and late positions:

```powershell
$env:XLIDE_PERF_WORKBOOK = 'F:\GitHub\xlide\xlide_vscode_testing\ROneCOne_Delegates_Demo.xlsm'
npx vitest run tests/incrementalLexerWorkbookPerformance.test.ts
```

The real editor harness copies the workbook into its disposable test workspace:

```powershell
npm run compile
npx vscode-test --code-version 1.139.1 --grep 'Actual large class latency'
```

## Measurements

On the 969,296-character ROneCOne class, medians of 15 samples per position,
incremental versus full lex, on the same machine and in the same process:

| Edit position | Insert one character | Replace one character | Full lex |
| --- | ---: | ---: | ---: |
| 1% | 11.17 ms | 2.12 ms | 28.56–32.87 ms |
| 50% | 5.09 ms | 1.76 ms | 28.06–29.15 ms |
| 99% | 1.83 ms | 1.75 ms | 27.93–28.44 ms |

Early insertions still clone most suffix tokens because their absolute offsets
change. Equal-length replacements can reuse suffix tokens directly.

The real VS Code 1.139.1 harness with compatible-source selection recorded
edit-to-completion results of 162.65, 160.66, 93.66, 126.83 and 75.05 ms.
Fresh declaration edit to hover took 95.86 ms. These are individual observations,
not an isolated before/after comparison or a claim of instantaneous response.
Provider traces still showed 43–82 ms completion work and 62–82 ms semantic-token
work. A final run of all 17 editor checks recorded completion results of 166.44,
110.12, 78.29, 141.08 and 69.07 ms, and fresh-edit hover of 111.77 ms.
Full parsing and background work remain targets for issue #985.

Portable checks compare every insertion/deletion position in a lexical fixture
against the full lexer, including canonical keywords, UTF-16 offsets, physical
line/column coordinates and trivia. Sequential edits cover CR, LF, CRLF, Unicode,
continuations, comments, strings and unchanged-snapshot guarantees. Performance
checks assert bounded re-lex work rather than timing thresholds.
