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
- [ ] **(12) Activity-bar icon** — only if promoting XLIDE out of the file Explorer view.

## Performance / reliability

- [x] **(13) Cache `listModules`/`listSubs` per workbook** — `_modulesListCache` in `XlsmExplorer` avoids repeated bridge round-trips; cleared on `refresh()` so edits always re-fetch.
- [x] **(14) Debounce filesystem watcher** — coalesces explorer refreshes to 200 ms in `extension.ts`.
- [x] **(15) Cancellation tokens on RPC** — `bridge.call()` now accepts an optional `CancellationToken`; pending requests are rejected with `CancellationError` on cancellation.

## Live Share polish

- [ ] **(16) Pre-translate paths via `convertLocalUriToShared`** — embed `vsls:` paths instead of opaque workbookIds, so right-clicks could reach the host's shared file.
- [ ] **(17) Guest follow-on-open** — broadcast `xlide.openModule` invocations so guests can follow the host.

## Developer experience

- [ ] **(18) Smoke test command** — `xlide.dev.smoke` runs listModules / readModule / writeModule roundtrip against a checked-in fixture.
- [ ] **(19) TS unit tests** — `tests/` folder with Mocha or vitest for pure-logic (URI encode/decode, sidecar JSON, managedFiles diff).
- [ ] **(20) CI workflow** — `.github/workflows/ci.yml` running `npm run compile` + `pytest`.
- [ ] **(21) CHANGELOG.md** + early releases tagged as `--pre-release` on the marketplace.

## Files to Keep Up To Date

- `docs/architecture.md` — structure changes
- `docs/roadmap.md` — this file
- `README.md` — dev-facing docs
- `MARKETPLACE.md` — user-facing marketplace listing
- `CHANGELOG.md` — when added
