# XLIDE - VBA for VS Code

Edit Excel VBA projects directly in VS Code. XLIDE turns macro workbooks into
first-class development projects: browse workbook modules, edit writable
`xlide-vba` documents, use Smart Enter and IntelliSense, run deterministic
analysis, sync modules with files, execute macros and workbook tests through
Excel when needed, and give AI agents workbook-aware tools.

<a href="https://github.com/sponsors/WilliamSmithEdward"><img src="https://img.shields.io/badge/Sponsor-%E2%9D%A4-pink?style=for-the-badge" alt="Sponsor WilliamSmithEdward"></a>

---

## Highlights of What's New In Version 2

Version 2 turns XLIDE from a workbook module browser into a fuller VBA
development environment for VS Code:

- Dedicated XLIDE Activity Bar/sidebar with setup health, selected-workbook
  actions, global settings, and support commands.
- Expanded deterministic VBA language service: Smart Enter, block snippets,
  completions, hover, signature help, semantic type coloring, navigation,
  rename, and focused code actions.
- Live, current-module, and workbook-wide analysis with stable rule codes,
  severity controls, suppressions, workbook/global settings, and a dedicated
  analysis results panel.
- Stronger project-aware understanding for visible identifiers, type names,
  source-backed class members, UDT fields, known runtime signatures, and the
  first generated Excel host-member surfaces.
- Previewable import/export sync for workbook modules, backed by the current
  `<workbook>.xlide_settings.json` sidecar model.
- Workbook test runner for marked `@xlide-test` procedures, including
  `XlideAssert`, read-only Excel test execution, artifacts, and CI status output.
- AI-agent tools for workbook discovery, VBA reads/writes, analysis, tests,
  sheet/cell I/O, formula reads, module export, and openpyxl automation.
- Safety and support hardening: explicit workbook mutation, write summaries,
  settings validation, support bundles, setup diagnostics, and recovery notes.

---

## Start Here

If you are new to using XLIDE, start with the user guides:

| Need | Guide |
|---|---|
| First setup and first workbook workflow | [Getting started](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/getting_started.md) |
| Diagnostics, analysis results, and ignores | [Analysis and ignores](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/analysis.md) |
| Import/export and module sync | [Import and export](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/sync.md) |
| Writing and running workbook tests | [Testing VBA workbooks](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/testing.md) |
| AI-agent and CI workflows | [Automation and CI](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/automation.md) |
| Trust model, Excel setup, support bundles, and recovery | [Safety, trust, and support](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/support.md) |
| Guide index | [User guides README](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/README.md) |

The sections below are primarily for extension contributors and packagers.

---

## What XLIDE Does

- Opens `.xlsm`, `.xlsb`, and `.xlam` VBA modules as editable VS Code documents.
- Saves module edits back to the workbook with normal VS Code save behavior.
- Provides VBA syntax highlighting, Smart Enter, block snippets, completions,
  hover, signature help, Go to Definition, Find All References, Rename Symbol,
  semantic type coloring, and deterministic code actions where XLIDE can prove
  the edit.
- Runs live diagnostics, current-module analysis, and workbook-wide analysis
  with stable rule codes, severity controls, suppressions, and a dedicated
  analysis results panel.
- Imports and exports workbook modules through previewable sync GUIs so `.bas`
  and `.cls` files can be reviewed or source-controlled.
- Runs macros and marked workbook tests through explicit Excel COM workflows on
  Windows.
- Exposes workbook discovery, module editing, analysis, tests, sheet/cell I/O,
  formula reads, module export, and openpyxl automation as VS Code language
  model tools.
- Keeps setup health, global settings, workbook actions, and support commands
  discoverable from the XLIDE Activity Bar view.

---

## Requirements

- **VS Code** 1.95+
- **Python 3.10+** -- the VBA read/write backend runs as a child process
- Python packages: `pyOpenVBA >= 3.0.1`, `openpyxl >= 3.1.0`

Reading, editing, analyzing, importing, exporting, and sheet/cell inspection use
the Python backend and do not require Excel COM automation.

Microsoft Excel COM is required for workflows that execute VBA, such as running
a macro or running workbook tests. Those execution workflows are Windows-only
and require the workbook's macro security state to allow execution.

---

## Development setup

```bash
git clone https://github.com/WilliamSmithEdward/xlide_vscode.git
cd xlide_vscode

# TypeScript side
npm install
npm run compile        # type-check + esbuild bundle -> out/extension.js

# Python side (optional venv)
python -m venv .venv
.venv\Scripts\activate   # or: source .venv/bin/activate
pip install -r python/requirements.txt
```

Press **F5** in VS Code to launch an Extension Development Host with the
extension loaded and the watch compiler running.

---

## Repository Map

```
xlide_vscode/
  src/
    extension.ts              # Activation entry point
    pythonBridge.ts           # JSON-RPC client for the Python backend
    xlideFileSystem.ts        # xlide-vba:// writable virtual filesystem
    xlsmExplorer.ts           # Explorer-hosted workbook/module tree
    xlideSidebar.ts           # XLIDE Activity Bar/sidebar webview
    commands.ts               # Workbook, module, analysis, sync, run/test commands
    agentTools.ts             # VS Code language model tool registrations
    workbookSettings.ts       # Strict <workbook>.xlide_settings.json owner
    vbaLanguageProviders.ts   # Editor providers, diagnostics, Smart Enter
    analyzer/                 # Lexer, parser, binder, diagnostics, completion
  python/
    server.py                 # JSON-RPC 2.0 server
    xlide/
      vba_io.py               # pyOpenVBA workbook/module operations
      excel_io.py             # openpyxl sheet/cell/formula operations
  syntaxes/
    vba.tmLanguage.json       # TextMate grammar
  language-configuration/
    vba-language-configuration.json
  user_guides/                # Public user guides
  docs/
    architecture.md           # Full architecture reference
    roadmap_version_2.x.md    # Closed v2 launch-hardening roadmap
    roadmap_version_3.x.md    # Forward backlog
```

### Key design decisions

| Decision | Rationale |
|---|---|
| Long-lived Python process | Amortizes Python startup across workbook requests |
| `FileSystemProvider` over `TextDocumentContentProvider` | Read/write virtual FS: Ctrl+S triggers `writeFile` with no custom save command |
| Virtual URI scheme `xlide-vba://` | Decouples workbook path + module name from the editor's file concept |
| Shared workbook settings owner | Global settings stay in VS Code; workbook overrides live beside the workbook |
| Python for workbook I/O, COM only for execution | Basic editing and analysis stay portable; macro/test execution uses Excel COM explicitly |
| Confirmation on write tools | Prevents AI agents from silently mutating production workbooks where VS Code supports confirmation |

### VBA language ID

Registered as `vba` in `package.json` with extensions `.bas`, `.cls`, `.frm`.
The TextMate grammar in `syntaxes/vba.tmLanguage.json` is scoped to
`source.vba` and covers all reserved identifiers from MS-VBAL v20250520
(section 3.3.5.2: statement-keywords, marker-keywords, operator-identifiers,
reserved-names, special-forms, reserved-type-identifiers, literal-identifiers,
def-type directives, and implementation-reserved identifiers).

See [docs/architecture.md](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/docs/architecture.md) for the full architecture,
Python bridge method list, workbook settings schema, import/export planning,
and agent-tool implementation notes.

---

## Build commands

| Command | Purpose |
|---|---|
| `npm run compile` | Type-check + dev bundle |
| `npm run watch` | Incremental type-check + esbuild watch |
| `npm test` | Run the Vitest test suite |
| `npm run package` | Production bundle (minified) |
| `npm run vsix` | Build a versioned `.vsix` under `dist/` |
| `npm run test:oracle:vbe` | Optional Excel/VBE oracle checks for language-behavior evidence |

---

## AI Agent Tools

XLIDE contributes VS Code language model tools so Copilot and compatible agents
can inspect, edit, analyze, test, and sync workbooks without guessing where VBA
source lives. The workbook and `xlide-vba` virtual modules are the source of
truth; exported `.bas` and `.cls` files are sync artifacts unless the user asks
to work with them directly.

Recommended flow:

1. Discover the workbook with `xlide_listWorkbooks` or `xlide_getWorkbookInfo`.
2. Read workbook VBA with `xlide_readModule`.
3. Write workbook VBA with `xlide_writeModule` when an edit is approved.
4. Run `xlide_analyzeWorkbook`.
5. Run `xlide_runVbaTests` when Excel COM is available and workbook tests exist.

| Tool family | Tools |
|---|---|
| Workbook discovery and validation | `xlide_listWorkbooks`, `xlide_getWorkbookInfo`, `xlide_validateWorkbook` |
| VBA module discovery and editing | `xlide_listModules`, `xlide_listSubs`, `xlide_readModule`, `xlide_writeModule`, `xlide_renameModule`, `xlide_deleteModule` |
| Analysis and test execution | `xlide_analyzeWorkbook`, `xlide_runVbaTests` |
| Workbook and worksheet data | `xlide_createWorkbook`, `xlide_listSheets`, `xlide_readCells`, `xlide_readFormulas`, `xlide_writeCells`, `xlide_runOpenpyxl` |
| Module export/sync | `xlide_exportModules`, `xlide_configureExportMode` |

Most workbook-mutating agent tools request explicit VS Code confirmation before
writing modules, cells, files, or workbook settings. See the
[Automation and CI guide](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/automation.md) for the supported
agent/CI workflow.

---

## Workbook Settings

Workbook-specific settings are stored beside the workbook as
`<workbook>.xlide_settings.json`, for example
`Budget.xlsm.xlide_settings.json`. Prefer the XLIDE import/export, analysis,
test, and settings GUIs for normal editing; they preserve unrelated settings and
show the effective source of each value.

```json
{
  "exportFolder": "C:/absolute/path/to/export",
  "exportMode": "exportAll",
  "importMode": "updateOnly",
  "analysis": {
    "visibleSeverities": ["error", "warning", "information"],
    "untrackedRules": [],
    "ruleSeverityOverrides": {}
  },
  "tests": {
    "artifactFolder": "tests",
    "artifactRetention": 20
  }
}
```

Malformed workbook settings are reported as explicit XLIDE settings errors
instead of being ignored. Global defaults live in VS Code machine/profile
settings; workbook sidecars contain only workbook-specific overrides.

---

## Live Share

XLIDE VBA browsing for Live Share **guests** is currently not supported.
Microsoft's Live Share platform restricts the shared-service RPC channel
(`vsls.shareService`) to extensions on a curated first-party allowlist, so
third-party extensions like XLIDE cannot proxy VBA read/write calls from a
guest to the host. The XLIDE Explorer therefore returns an empty tree for
guests and shows an informational welcome view.

What still works in a Live Share session:

| Role | XLIDE behavior |
|---|---|
| **Host** | Full local VBA editing -- open, edit, save `.xlsm`/`.xlsb`/`.xlam` modules exactly as if no session were active. |
| **Guest** | Can fully view and edit any VBA module the host has open in the editor (Live Share shares those buffers normally). Cannot browse the XLIDE Explorer or open new modules independently -- only the host can navigate and open them. XLIDE panel shows a "not supported" notice. |
| **Guest without XLIDE installed** | No action needed -- XLIDE is host-only. Joining a session does not require the extension. |

Related upstream issue: [microsoft/live-share#4877](https://github.com/microsoft/live-share/issues/4877)
(third-party `shareService` allowlist, closed as Not Planned).

---

## Further reading

- [User guides README](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/README.md) -- public guide index
- [Getting started guide](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/getting_started.md) -- first-hour XLIDE workflow
- [Analysis and ignores guide](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/analysis.md) -- diagnostics, settings, suppressions, and tracking
- [Import and export guide](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/sync.md) -- previewable `.bas`/`.cls` module sync
- [Testing guide](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/testing.md) -- workbook test authoring and execution
- [Automation guide](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/automation.md) -- agent and CI workflows
- [Safety and support guide](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/support.md) -- trust model, Excel setup, and recovery
- [docs/architecture.md](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/docs/architecture.md) -- full architecture reference
- [MS-VBAL specification](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/)
- [pyOpenVBA](https://github.com/WilliamSmithEdward/pyOpenVBA)
