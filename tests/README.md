# Tests

Run commands from the repository root. `npm test` discovers all
`tests/**/*.test.ts` files, including nested feature folders.

```sh
npm test
npx vitest run tests/navigation/ tests/projectManagement/
npx vitest run tests/moduleSync/
npm run check-types
```

Use a trailing slash to select a folder: Vitest filters by path substring,
and the slash keeps the selection specific to that directory. A filename
selects an individual suite.

## Where tests belong

Group by the behavior under test, rather than every implementation module it
imports. A diagnostic regression that uses the parser belongs with diagnostics;
completion behavior belongs with providers. Keep performance and work-count
regressions beside the feature they measure.

- `access/`: Access storage, VBA writes, container ordering, design metadata and themes.
- `agents/`: Agent tools, instructions, review decorations, API server, edit mirroring and write audit.
- `annotations/`: Attribute/folder annotation parsing, rewriting, source lines and duplicates.
- `application/`: Sidebar, command registration/logs, output logs and support bundles.
- `codeActions/`: Diagnostic quick fixes, declaration insertion and suppression actions.
- `conditional/`: Conditional compilation, compiler/project constants, replay and inactive arms.
- `constants/`: Integer literals, constant expressions, dependency resolution and stack safety.
- `containers/`: CFB/ZIP/OVBA codecs, macro containers, project creation and byte preservation.
- `diagnostics/`: Diagnostic rules, shared inference/dataflow, suppressions and regression/work-count checks.
- `docs/`: Documentation comments, external XML metadata, recovery and tag source mapping.
- `editor/`: Typing commands, Enter/Tab/Backspace, canonical casing and document lifecycle.
- `expressions/`: Expression type resolution and enclosing-procedure lookup.
- `fixtures/`: Shared checked-in inputs and fixture provenance notes.
- `flow/`: Procedure labels and unstructured control-flow analysis.
- `formatting/`: Formatter behavior, commands, token reuse and performance regressions.
- `forms/`: Office UserForm/control storage, editing, preview, launch, markup and round trips.
- `git/`: Module history, comparison and change decorations.
- `helpers/`: Shared utilities/mocks, including fixture loaders.
- `hostModels/`: Office object models, member/type metadata, constants and interop contracts.
- `language/`: Code pages, Unicode names/combining marks and cross-platform language checks.
- `lexer/`: Tokenization, keywords, incremental lexing and token lookup/cache work.
- `libraryReferences/`: Project library references, dependencies, removal commands and qualifiers.
- `moduleSync/`: Module export, import/export plans, previews and synchronization round trips.
- `navigation/`: Symbol references/definitions, renames/history, caret and procedure lookup.
- `office/`: Live Office suites and harness helpers; uses the separate Office configuration.
- `officeHost/`: Host launch, COM availability and view placement unit tests.
- `package/`: Manifest, security scanner and workflow-pinning checks.
- `parser/`: Parsing, incremental reuse and structural/source-scan equivalence.
- `platform/`: Node/browser shims, host platforms, PowerShell, locks and low-level utilities.
- `projectAnalysis/`: Project/shared/incremental/live analysis, symbol indexes, contribution reuse and workers.
- `projectExplorer/`: Project tree, folders, follow/readiness, refresh, sorting and lifetime.
- `projectManagement/`: Project identity/selection, files/modules, URI/tabs, protection, locks and coordinated writes.
- `properties/`: Property-based checks.
- `providers/`: Completion, hover, signatures, semantic tokens and editor-query work-count regressions.
- `refactor/`: Refactorings, call sites, shared edit helpers and command bindings.
- `runtime/`: VBA built-in metadata, documentation and reference coverage.
- `settings/`: Global/project settings, manifests and settings webviews.
- `shapes/`: Shape editing/preview, tree rows and host-specific shape surfaces.
- `source/`: Source-line spans, module headers/members, identifier occurrences and positions.
- `vb6/`: VB6 projects, designers, association/location, code pages and object models.
- `vbaTest/`: VBA test runner, execution pipeline, artifacts, support modules and results UI.
- `worksheets/`: Worksheet storage, controls, cells, formulas and sheet-part reads.

The remaining root `*.test.ts` files cover active oracle/corpus work
(`oracle*`, `vbeOracle*`, `twinbasic*` and `corpusProvenance`). They stay in place
while that work is in flight. Do not add unrelated feature suites to the root.

## Moving or adding tests

Prefer an existing feature folder. When moving a suite, preserve assertions
and fixtures; adjust imports, mock targets, file-relative assets, and explicit
CI/documentation paths. Shared inputs stay in `fixtures/`. Verify discovery
and run the affected suites. Keep behavior changes separate from moves.

Performance/work notes live in `docs/`; fixture-specific provenance belongs
beside the fixture. Update this guide when adding a feature folder.

## Other harnesses

- `npm run test:integration`: compiles and runs the VS Code tests in `src/test/`.
- `npm run test:office`: runs `tests/office/**/*.office.ts` with the separate Office config. Requires Windows and usable Office installations. The harness skips hosts it cannot safely use; supported hosts are Excel, Word, PowerPoint and Access.
- `npm run test:oracle:vbe`: runs the VBE oracle; see the [oracle setup](../syntax_corpus/oracle/README.md).
- `npm run test:oracle:twinbasic`: runs the twinBASIC oracle; see its [setup guide](../syntax_corpus/oracle/twinbasic/README.md).

Some unit benchmarks require an explicitly configured workbook or environment
variable and otherwise skip. Follow their reproduction notes in `docs/`.
Passing unit tests does not measure live Office or editor painting latency.
