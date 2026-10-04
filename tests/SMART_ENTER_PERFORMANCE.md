# Smart Enter performance checks

Run deterministic work-count tests and the lexer substrate corpus comparison:

```powershell
npx vitest run tests/smartEnterSurfacePerformance.test.ts tests/smartEnterSubstrateComparison.test.ts tests/vbaStructuralAnalysis.test.ts
```

The listener tests assert zero full-document reads after ordinary Enter and
whole-line comment Enter. Leading-dot continuation and block openers still read
the module when they need to find an active With or a matching closer.
The stripping tests check UTF-16 columns and continued comments across LF, CRLF,
and CR; the corpus comparison covers the repository's VBA samples.

Run the opt-in component benchmark on an otherwise idle machine:

```powershell
New-Item -ItemType Directory -Force out | Out-Null
$env:XLIDE_ENTER_BENCHMARK_OUTPUT = 'out/smart-enter-latency.json'
npx vitest run tests/smartEnterSurfacePerformance.test.ts -t 'measures Enter and stripped-line work'
Remove-Item Env:XLIDE_ENTER_BENCHMARK_OUTPUT
Get-Content out/smart-enter-latency.json
```

It measures 21 samples per median after warm-up with a 1,200-procedure, 78,179-byte
module containing strings and comments. VS Code is mocked and the text getter
joins an array of lines. This measures the Smart Enter listener and shared
stripped-line helper, excluding other typing listeners and renderer latency.

Exercise real editor events and completion after an auto-inserted With dot:

```powershell
npm run compile
npx vscode-test --code-version 1.139.1 --grep 'Smart Enter surfaces|Canonical casing'
```

The integration harness copies the fixture workbook into its disposable workspace.
It verifies ordinary Enter in a large module, comment prefix continuation, With
dot insertion plus member completion, and missing block closer insertion. It
prints edit-command timings without machine-dependent pass thresholds; those
timings include command dispatch and other listeners, and exclude screen repaint.
