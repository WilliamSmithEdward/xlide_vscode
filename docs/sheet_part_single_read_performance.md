# Reading a sheet part once to list its shapes

Listing a workbook's modules gives each worksheet module its sheet's ActiveX controls as members (#225). To do that, `withSheetControls` calls `XlsxWorkbook.shapes()`, which inflated and decoded every worksheet part three times per call: once in `sheetCodeName`, once in `sheetParts` and once in `readControls`. A sheet part holds the sheet's rows, so `listModules`, `readModules` and `listShapes` cost in proportion to the data on the sheets, and the cached project parse did not help: the work repeated on every warm call. In the web build the pure-TypeScript inflate made each read several times dearer.

Two changes, both in `src/vba`:

1. `readSheet` reads the sheet part once and hands the text to `sheetParts` and `readControls`.
2. `sheetCodeName` looks for the `sheetPr` in the head of the part, through `ZipArchive.readPrefix`, before reading the whole of it. The `sheetPr` is the first child of the root element, so the head holds it. The parsing is the same two regular expressions as before. The first `sheetPr` of the head is the first of the part, so the answer is the same; a part whose head shows no `sheetPr` is read whole, as it always was.

That leaves one whole read of each sheet per listing, where there were three. A sheet with no `sheetPr` at all is read twice, once to look for one; Excel writes a `sheetPr` for every sheet that has a module, so in a workbook with a VBA project that is a sheet added and never seen by the VBA editor.

No cache was added, so there is nothing to invalidate. Nothing here runs per keystroke: the engine's `listModules` and `readModules` are what the Project Explorer calls to list a workbook's modules, what the symbol index and project-wide analysis call to read a project's source, and what module export, git compare, the form preview, the test runner and the agent tools call before they act.

## What can differ

Nothing is intended to, and nothing did. `readModulesFromBuffer` and `xlsx.shapes()` were run through the baseline sources and through these on all 28 Excel files under `tests/fixtures` and `assets/templates`, and on variants of each built for the purpose: every sheet, relationships, drawing and control part given an unsupported compression method, garbled from its start, its middle and near its end, deleted and emptied; each sheet grown past the head; its `sheetPr` removed, emptied, recased, duplicated across sheets, moved after the rows, and written with an entity, spaces around `=`, single quotes, a repeated attribute and a namespace prefix; and an ActiveX control entry added with and without the relationship that makes it one. That is 3,794 cases through the zlib codec and the web build's. The answers, error messages included, were the same in every one.

## Reproduce

Node v24.21.0; Intel Core i5-8265U @ 1.60GHz (a laptop, and noisy: read the ratios, not the absolute values). Baseline `982d124cc98b0b0ade198763365aa5d3aff22dda`. The script bundles `projectService` with the pinned esbuild, once for Node and once with the web build's leaf swap, and with `--baseline` loads the two changed files from that commit. Each fixture's first sheet is grown to 60,000 non-repeating rows (5.5 MB of XML). Each call is warm: the project parse is cached, three warm-up calls, then 15 measured rounds. Timing covers the engine call only, not VS Code or the tree. The web rows are the web build's bundle run under Node, not in a browser.

```powershell
node scripts/benchmark-module-listing.mjs --baseline=982d124cc98b0b0ade198763365aa5d3aff22dda > before.json
node scripts/benchmark-module-listing.mjs > after.json
```

Every row carries a digest of what `listModules`, `readModules` and `listShapes` returned; the digests matched before and after on every row.

## Results

Milliseconds per call. `SheetsFixture`'s large sheet has no shapes; `ShapesFixture`'s has a drawing and form controls.

| Workbook | Codec | Call | Median before | Median after | p95 before | p95 after |
| --- | --- | --- | ---: | ---: | ---: | ---: |
| SheetsFixture | node | listModules | 32.9 | 14.4 | 46.5 | 24.1 |
| SheetsFixture | node | readModules | 38.5 | 13.9 | 48.9 | 21.9 |
| SheetsFixture | node | listShapes | 36.8 | 13.8 | 50.2 | 22.6 |
| SheetsFixture | web | listModules | 139.2 | 37.0 | 203.5 | 47.0 |
| SheetsFixture | web | readModules | 101.5 | 36.9 | 111.5 | 47.4 |
| SheetsFixture | web | listShapes | 100.1 | 38.0 | 108.2 | 48.9 |
| ShapesFixture | node | listModules | 31.4 | 13.9 | 46.0 | 22.1 |
| ShapesFixture | node | readModules | 31.1 | 13.5 | 44.4 | 23.2 |
| ShapesFixture | node | listShapes | 31.5 | 14.4 | 43.1 | 25.4 |
| ShapesFixture | web | listModules | 99.7 | 45.2 | 115.9 | 49.7 |
| ShapesFixture | web | readModules | 115.8 | 44.8 | 136.6 | 63.3 |
| ShapesFixture | web | listShapes | 101.6 | 51.5 | 113.6 | 69.5 |

With `--rows=0` (the fixtures as they are, a few kilobytes per sheet) every call stayed between 1.0 and 2.4 ms before and after; the differences are within the noise, so small workbooks neither gain nor lose.

## Validation

`tests/sheetPartReads.test.ts` counts whole-part reads rather than time: one for a large sheet when its shapes are listed and when the workbook's modules are, two for a large sheet with no `sheetPr`. Against the baseline sources those three tests fail (the counts were three). The same file checks the code name `shapes()` gives against the first `sheetPr` of the whole part, for every `.xlsm`, `.xlam` and `.xltm` in the repository and for a large sheet whose `sheetPr` is written nine ways, one of them after the rows where the head does not reach.
