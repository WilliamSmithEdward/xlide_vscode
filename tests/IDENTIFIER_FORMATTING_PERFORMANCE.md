# Identifier candidate formatting performance

The identifier resolver filters names and applies shadowing before formatting
candidate details and documentation. Returned completion fields remain strings,
and metadata is read again on each request.

Run the work-count and behavior regressions:

```powershell
npx vitest run tests/identifierFormattingPerformance.test.ts tests/vbaIdentifierCompletion.test.ts tests/vbaCanonicalCaseSymbolReuse.test.ts
```

Run the opt-in benchmark on an otherwise idle machine:

```powershell
New-Item -ItemType Directory -Force out | Out-Null
$env:XLIDE_IDENTIFIER_BENCHMARK_OUTPUT = 'out/identifier-formatting.json'
npx vitest run tests/identifierFormattingPerformance.test.ts -t 'measures prefix-filtered identifier and casing work'
Remove-Item Env:XLIDE_IDENTIFIER_BENCHMARK_OUTPUT
Get-Content out/identifier-formatting.json
```

The report uses 21 warmed samples for a 103,345-byte module with 1,201
procedures and 1,200 additional project procedures. It measures a
narrow identifier prefix, casing one body line, and a separate unfiltered list
request. Lexer/parser caches warm normally. These are pure analyzer CPU timings,
excluding project I/O, VS Code command dispatch, and renderer latency.

Exercise actual editor completion and typing:

```powershell
npm run compile
npx vscode-test --code-version 1.139.1 --grep 'Identifier surfaces|Canonical casing|Editor surfaces'
```

The integration harness uses its disposable fixture workbook. It checks
prefix-filtered project documentation in a large module, individual character
edits followed by idle casing, and local documentation winning over a project
duplicate, alongside existing typing, dot completion, hover and caret tests.
