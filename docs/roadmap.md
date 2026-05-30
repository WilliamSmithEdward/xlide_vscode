# XLIDE Roadmap

Quality-of-life and polish backlog, ordered by **impact-per-effort**. Items already implemented are marked `[x]`. North star: match Visual Studio 2022 VB editing experience as closely as possible.

## Top 5 (implemented in this pass)

- [x] **(1) Workbook-locked error UX** — detect "file in use by Excel" on `writeModule`/`readModule` and show a Retry action.
- [x] **(2) VBA snippets** — `snippets/vba.json` for `sub`, `func`, `for`, `forEach`, `with`, `select`, `class`, `prop`.
- [x] **(3) Bridge auto-restart** — on unexpected Python child exit, mark the bridge as stopped so the next `call()` rejects cleanly with a clear, actionable error instead of hanging.
- [x] **(4) VBA `onEnterRules` / auto-pair** — auto-insert `End Sub`/`End Function`/`End If`/`Next`/`Loop` etc. like the VBA IDE.
- [x] **(5) Status bar items** — show `XLIDE: <workbook>` on the active module, plus `XLIDE (Live Share): N` for guests.

## High-leverage UX

- [x] **(6) Reveal in XLIDE Explorer** — `treeView.reveal(node)` for the active module.
- [x] **(7) Auto-expand sidebar on first .xlsm** — expand the first workbook automatically on activation.
- [x] **(8) Persist last-opened modules** — handled automatically by VS Code's editor restoration via the `onFileSystem:xlide-vba` activation event; no additional code required.
- [x] **(9) Welcome notification on first activation** — single, dismissible nudge replacing the deleted walkthrough.

## Editor polish

- [x] **(10) Document outline icons** — symbol-kind mapping extended to cover Sub/Function/Property/Const/Type/Enum in `vbaSymbolIndex.ts` and `vbaLanguageProviders.ts`.
- [x] **(11) Command palette categorization audit** — every `xlide.*` command declares `"category": "XLIDE"`.
- [ ] ~~**(12) Activity-bar icon**~~ — won't implement for now; XLIDE stays in the File Explorer view.

## Performance / reliability

- [x] **(13) Cache `listModules`/`listSubs` per workbook** — `_modulesListCache` in `XlsmExplorer` avoids repeated bridge round-trips; cleared on `refresh()` so edits always re-fetch.
- [x] **(14) Debounce filesystem watcher** — coalesces explorer refreshes to 200 ms in `extension.ts`.
- [x] **(15) Cancellation tokens on RPC** — `bridge.call()` now accepts an optional `CancellationToken`; pending requests are rejected with `CancellationError` on cancellation.

## Live Share polish

- [ ] ~~**(16) Pre-translate paths via `convertLocalUriToShared`**~~ — won't implement for now; Live Share API is opaque and poorly documented.
- [ ] ~~**(17) Guest follow-on-open)**~~ — won't implement for now; depends on Live Share coordination layer.

## Developer experience

- [x] **(18) Smoke test command** — `xlide.dev.smoke` (XLIDE: Run Smoke Test) — finds a workbook in the workspace, runs `listModules` + `readModule`, and reports results in the XLIDE Output channel.
- [x] **(19) TS unit tests** — `tests/vbaParsing.test.ts` (parseVbaModule: Sub/Function/Property/Const/Enum/Type, spans, visibility) and `tests/uriCodec.test.ts` (decodeModuleUri: module name, URL-encoding, extension variants, error cases). Run with `npm test` (vitest).
- [x] **(20) CI workflow** — `.github/workflows/ci.yml` runs `npm run compile` + `npm test` (TypeScript job) and `pytest python/tests/` (Python job) on push and pull requests.
- [x] **(21) CHANGELOG.md** — created at repo root; all implemented features documented.

## Protected, signed & lifecycle

- [x] **(22) Protected-workbook editing** — `writeModule`/`renameModule`/`deleteModule` save with `allow_protected=True` so password-locked VBA projects edit in place.
- [x] **(23) Signature-invalidation notice** — dropped digital signatures are detected (`signatureDropped`) and surfaced once per workbook via `notifySignatureDropped`, instead of being silenced.
- [x] **(24) Protection/signature badges** — `XlsmExplorer` lazily probes `getProtectionInfo` and shows `[locked]`/`[signed]` tags on workbook nodes.
- [x] **(25) Validate VBA Project** — `xlide.validateWorkbook` command + `xlide_validateWorkbook` tool wrap `ExcelFile.validate()`, reporting issues to the Output channel.
- [x] **(26) New Macro-Enabled Workbook** — `xlide.newWorkbook` command + `xlide_createWorkbook` tool scaffold a fresh `.xlsm`/`.xlsb` via `ExcelFile.create_new`.

## Files to Keep Up To Date

- `docs/architecture.md` — structure changes
- `docs/roadmap.md` — this file
- `README.md` — dev-facing docs and marketplace listing
- `CHANGELOG.md` — when added
