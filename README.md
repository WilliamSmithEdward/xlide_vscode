# XLIDE - Excel VBA for VS Code

XLIDE is a VS Code extension that brings Excel .xlsm macro files into the editor
as a first-class development experience - tree view, editable VBA modules, and
full AI agent tool integration (GitHub Copilot).

## Features

- **VBA Explorer** - sidebar tree showing every .xlsm/.xlsb/.xlam in the workspace,
  expanded to modules, then to individual Sub/Function/Property procedures.
- **Edit modules in VS Code** - click any module to open its VBA source. Press
  Ctrl+S (or Cmd+S) and the change is written back into the .xlsm automatically.
  No Office installation required.
- **Add / Rename / Delete modules** - right-click any entry in the explorer.
- **Read Excel cell data** - via openpyxl (formula values resolved).
- **AI agent tools** - six tools registered with VS Code's Language Model API so
  Copilot (and other LM agents) can list, read, and write VBA modules and cell
  data directly.

## Requirements

- VS Code 1.95 or newer
- Python 3.10 or newer with the following packages installed:

```
pip install pyOpenVBA openpyxl
```

  Or using the bundled requirements file:

```
pip install -r python/requirements.txt
```

- Node.js (for building the extension from source)

## Setup

### Install Python dependencies

```bash
pip install -r python/requirements.txt
```

### Build the extension

```bash
npm install
npm run compile
```

### Run / debug

Press **F5** in VS Code to launch the Extension Development Host.

## Configuration

| Setting | Default | Description |
|---|---|---|
| `xlide.pythonPath` | `""` | Path to the Python executable. Leave empty to use `python` (Windows) or `python3` (Mac/Linux) from PATH. |

## AI Agent Tools

The following tools are available to Copilot and other VS Code LM agents:

| Tool reference | Description |
|---|---|
| `#xlideListModules` | List all VBA modules in a workbook |
| `#xlideListSubs` | List all Sub/Function procedures in a module |
| `#xlideReadModule` | Read full VBA source of a module |
| `#xlideWriteModule` | Write (overwrite) a VBA module and save the workbook |
| `#xlideReadCells` | Read a cell range from a worksheet |
| `#xlideWriteCells` | Write values to a cell range and save the workbook |

Write operations require user confirmation in the Copilot chat UI.

## Architecture

See [docs/architecture.md](docs/architecture.md) for the full architecture overview.

## License

MIT
