# Tests

Run the unit suite from the repository root with `npm test`. Vitest discovers
`tests/**/*.test.ts`, including nested suites. To run one area:

```sh
npx vitest run tests/parser tests/lexer tests/providers
npx vitest run tests/access
npx vitest run tests/docs/ tests/annotations/
npx vitest run tests/conditional/ tests/constants/
npm run check-types
```

## Where tests belong

- `constants/`: integer literal parsing, constant-expression evaluation, dependency resolution, and stack-safety regressions.
- `conditional/`: conditional compilation, compiler/project constants, constant replay, inactive arms, lookup reuse, and expression stack safety. Diagnostic-rule tests stay in `diagnostics/`.
- `annotations/`: attribute annotation parsing/rewriting, duplicates, physical lines, and folder annotation parsing/equivalence.
- `docs/`: VBA documentation comments, XML metadata parsing/recovery, tag source mapping, and documentation diagnostics. Word document shape tests belong with document/container behavior instead.
- `access/`: Access database storage, VBA writes, container ordering, design metadata, and themes; shared inputs stay in `fixtures/`.
- `parser/`: expression and module parsing, incremental parse reuse, and parser cache pressure.
- `lexer/`: tokenization, keyword classification, incremental lexing, and token lookup/cache work.
- `providers/`: editor completion, hover, signature help, navigation, colors, semantic tokens, and scheduling/latency regressions. Tests of the analyzer helpers behind those surfaces belong here too when the editor behavior is their main subject.
- `projectExplorer/`: project tree, refresh, readiness, sorting, and lifetime regressions.
- `vbaTest/`: the VBA test runner, support modules, artifacts, and results UI.
- `formatting/`: formatter behavior, commands, and formatting performance.
- `settings/`: global and project settings, manifests, and settings webviews.
- `diagnostics/`: diagnostic rules and their regressions.
- `properties/`: property-based checks.
- `office/`: the live Office harness and suites; see the separate command below.
- `helpers/`: shared test utilities, including the VS Code mock.
- `fixtures/`: checked-in inputs shared by tests.

Other feature suites still live at the root while organization proceeds. Prefer
an existing feature folder when one fits. Group tests by the behavior under
test rather than every implementation module they import; a diagnostic test
that uses the parser still belongs with diagnostics. Keep performance and
work-count regressions beside their feature's behavior tests.

When moving a suite, preserve its assertions and fixtures, adjust relative
imports and mock paths, and update explicit paths in CI and reproduction notes.
Check that recursive discovery still finds every test. Avoid folding behavior
changes into an organizational move.

## Other harnesses

- `npm run test:integration`: compiles and runs the VS Code integration tests in `src/test/`.
- `npm run test:office`: runs `tests/office/**/*.office.ts` using the separate Office config. Requires Windows and usable Office installations; the harness skips hosts it cannot safely use. The supported hosts are Excel, Word, PowerPoint, and Access.
- `npm run test:oracle:vbe`: runs the VBE oracle; setup and options are in [`syntax_corpus/oracle/README.md`](../syntax_corpus/oracle/README.md).
- `npm run test:oracle:twinbasic`: runs the twinBASIC oracle; see [`syntax_corpus/oracle/twinbasic/README.md`](../syntax_corpus/oracle/twinbasic/README.md).

Some unit benchmarks require an explicitly configured workbook or environment
variable and otherwise skip. Follow the relevant performance note to run them;
a normal unit run does not measure live Office or editor painting latency.
