# Tests

Four harnesses check XLIDE, each for what the others cannot reach.

| Harness | Command | What it runs |
| --- | --- | --- |
| Unit | `npm test` | Every `tests/**/*.test.ts` under Vitest, with `vscode` replaced by `tests/helpers/vscodeMock.ts`. CI runs it on Linux; this is the suite most changes need. |
| VS Code integration | `npm run test:integration` | `src/test/*.test.ts` inside a real VS Code with the extension loaded (`.vscode-test.mjs`). The `tests/*Integration.config.mjs` files are narrower runs of the same bundle with their own workspace and debugger port. |
| Live Office | `npm run test:office` | `tests/office/*.office.ts` against real Excel, Word, PowerPoint and Access (`vitest.office.config.ts`). Windows with Office only, not part of CI; a host it cannot use safely is skipped. |
| VBE oracle | `npm run test:oracle:vbe`, `npm run test:oracle:twinbasic` | `syntax_corpus/oracle/`: the corpus fixtures, each naming its `host`, compiled by the real VBE or twinBASIC to record what they accept. The unit suites in `tests/oracle/` check the runner and the recorded verdicts. |

## Where a suite lives

- `oracle/` - the VBE and twinBASIC oracle: the worker, startup, run lock, owned hosts, sound recovery, and the sweeps that hold the analyzer to the recorded verdicts.
- `engine/` - reading and writing Office and VB6 files: ZIP, compound file and OVBA containers, code pages, forms and shapes, Access storage and designs, module export, workbook parts.
- `agent/` - the agent tools, agent instructions and review, the MCP edit mirror, the API server and the write audit.
- `web/` - the browser build: the bundle and the web replacements for Node's buffer, crypto and path.
- `diagnostics/` - one suite per analyzer diagnostic or issue, each driven through `tests/helpers/diagnostics.ts`.
- `office/` - the live Office suite above.
- `properties/` - property-based checks; `.github/workflows/fuzz.yml` runs them at a higher count (`XLIDE_PROPERTY_RUNS`, `XLIDE_PROPERTY_SEED`).
- `helpers/` - shared setup: the `vscode` mock, fake config and project engine, temp folders, project fixtures and the oracle shard readers.
- `fixtures/` - Office files, Access databases, VB6 projects and recorded type data. Office extensions are gitignored, so a new binary fixture needs its `!` line in `.gitignore`.
- The root holds the suites not yet moved into a feature folder (#1441), and the checks of the repository itself: `packageManifest`, `workflowPins`, `securityScans` and `tempDirs`.

## Focused runs

```bash
npx vitest run tests/oracle
```

```bash
npx vitest run -t "ActiveSheet project class membership"
```

A path or folder narrows to those files; `-t` narrows to matching test names. Some CI jobs name their files in `.github/workflows/ci.yml`, so a moved or renamed suite is updated there too.

## Opt-in benchmarks

Workbook benchmarks are skipped unless their variable names a workbook, for example `XLIDE_PERF_WORKBOOK`, `XLIDE_ANALYSIS_PERF_WORKBOOK` or `XLIDE_PERF_MODULE`. Others are gated by a `XLIDE_*_BENCH` variable or write samples to a `XLIDE_*_BENCHMARK_OUTPUT` file. The suite itself names its variable. Benchmarks report typical and slow samples; the work-count guards beside them are what CI holds.
