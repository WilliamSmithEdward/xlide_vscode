# Reading a sheet part once to list its shapes

Listing a workbook's modules gives each worksheet module its sheet's ActiveX controls as members (#225). To do that, `withSheetControls` calls `XlsxWorkbook.shapes()`, which inflated and decoded every worksheet part three times per call: once in `sheetCodeName`, once in `sheetParts` and once in `readControls`. A sheet part holds the sheet's rows, so `listModules`, `readModules` and `listShapes` cost in proportion to the data on the sheets, and the cached project parse did not help: the work repeated on every warm call. In the web build the pure-TypeScript inflate made each read several times dearer.

Two changes, both in `src/vba`:

1. `readSheet` reads the sheet part once and hands the text to `sheetParts` and `readControls`.
2. `sheetCodeName` looks for the `sheetPr` in the head of the part, through `ZipArchive.readPrefix`, before reading the whole of it: the same head the sheet catalog reads. The `sheetPr` is the first child of the root element, so the head holds it. The parsing is the same two regular expressions as before. The first `sheetPr` of the head is the first of the part, so the answer is the same; a part whose head shows no `sheetPr` is read whole, as it always was.

That leaves one whole read of each sheet per listing, where there were three. A sheet with no `sheetPr` at all is read twice, once to look for one; Excel writes a `sheetPr` for every sheet that has a module, so in a workbook with a VBA project that is a sheet added and never seen by the VBA editor.

No cache was added, so there is nothing to invalidate. Nothing here runs per keystroke: the engine's `listModules` and `readModules` are what the Project Explorer calls to list a workbook's modules, what the symbol index and project-wide analysis call to read a project's source, and what module export, git compare, the form preview, the test runner and the agent tools call before they act.

## What can differ

Nothing is intended to. Reading the part once changes which call inflates it, not what is read from it. The code name is found by the same two regular expressions; the head of a part is its beginning, so the first `sheetPr` in the head is the first in the part, and a head with none sends the reader to the whole part as before. A part that cannot be read still fails, with the error it gave before.

The tests below hold the code name to that on every Excel fixture in the repository and on a large sheet written nine ways, and the benchmark's digests of what `listModules`, `readModules` and `listShapes` return match before and after. A wider one-off comparison of the baseline sources against these, over the fixtures and damaged variants of each through both codecs, found no difference either; its script is not part of this change.

## Reproduce

Node v20.20.2 on Windows; Intel Core i5-8265U @ 1.60GHz (a laptop, and noisy: read the ratios, not the absolute values). Baseline `751fdb2bcec07e2f4dc956d97bed5ae159bfb93a`. The script bundles `projectService` with the pinned esbuild, once for Node and once with the web build's leaf swap, and with `--baseline` loads `xlsx.ts` and `xlsxShapes.ts` from that commit. Each fixture's first sheet is grown to 60,000 non-repeating rows (5.5 MB of XML). Each call is warm: the project parse is cached, three warm-up calls, then 15 measured rounds. Timing covers the engine call only, not VS Code or the tree. The web rows are the web build's bundle run under Node, not in a browser.

```powershell
node scripts/benchmark-module-listing.mjs --baseline=751fdb2bcec07e2f4dc956d97bed5ae159bfb93a > before.json
node scripts/benchmark-module-listing.mjs > after.json
```

Every row carries a digest of what `listModules`, `readModules` and `listShapes` returned; the digests matched before and after on every row.

## Results

Milliseconds per call. `SheetsFixture`'s large sheet has no shapes; `ShapesFixture`'s has a drawing and form controls.

| Workbook | Codec | Call | Median before | Median after | p95 before | p95 after |
| --- | --- | --- | ---: | ---: | ---: | ---: |
| SheetsFixture | node | listModules | 35.4 | 15.1 | 71.7 | 16.6 |
| SheetsFixture | node | readModules | 44.2 | 15.1 | 60.1 | 16.9 |
| SheetsFixture | node | listShapes | 44.0 | 15.8 | 57.1 | 17.7 |
| SheetsFixture | web | listModules | 126.3 | 46.1 | 133.2 | 47.9 |
| SheetsFixture | web | readModules | 126.2 | 45.1 | 148.2 | 47.3 |
| SheetsFixture | web | listShapes | 126.5 | 44.9 | 134.3 | 49.4 |
| ShapesFixture | node | listModules | 33.3 | 16.4 | 35.2 | 27.5 |
| ShapesFixture | node | readModules | 30.0 | 16.2 | 34.5 | 17.8 |
| ShapesFixture | node | listShapes | 33.0 | 17.7 | 43.8 | 32.0 |
| ShapesFixture | web | listModules | 135.8 | 47.4 | 151.5 | 49.8 |
| ShapesFixture | web | readModules | 137.5 | 51.4 | 154.3 | 59.6 |
| ShapesFixture | web | listShapes | 125.4 | 47.3 | 130.7 | 57.8 |

With `--rows=0` (the fixtures as they are, a few kilobytes per sheet) every call stayed between 1.2 and 3.8 ms before and after; the differences are within the noise, so small workbooks neither gain nor lose.

## Validation

`tests/sheetPartReads.test.ts` counts whole-part reads rather than time: one for a large sheet when its shapes are listed and when the workbook's modules are, two for a large sheet with no `sheetPr`. Against the baseline sources those three tests fail (the counts were three). The same file checks the code name `shapes()` gives against the first `sheetPr` of the whole part, for every `.xlsm`, `.xlam` and `.xltm` in the repository and for a large sheet whose `sheetPr` is written nine ways, one of them after the rows where the head does not reach.
