# XLIDE - Excel VBA for VS Code

Browse and edit the VBA modules inside your `.xlsm` workbooks directly in
VS Code -- no round-tripping to the Excel IDE required.

---

## Requirements

- **Python 3.10+** ([download](https://www.python.org/downloads/)) -- add to PATH during install
- Python packages `pyOpenVBA` and `openpyxl` (XLIDE installs them for you on first run)

---

## Getting started

**1. Install Python 3.10+**
Download from [python.org](https://www.python.org/downloads/).
On Windows, tick **"Add Python to PATH"** on the first installer screen.

**2. Install the Python packages**
When you activate the extension, a prompt appears with an **Install Now**
button. Click it, or run this in any terminal:
```
pip install pyOpenVBA openpyxl
```

**3. Open a folder**
Use **File > Open Folder** and choose the folder that contains your `.xlsm`
files. The **XLIDE** panel appears in the Explorer sidebar.

**4. Edit a module**
Expand a workbook in the XLIDE panel, click any module, edit the code, and
press **Ctrl+S** to save back to the `.xlsm` file.

---

## Features

- **VBA Explorer** -- sidebar tree of every `.xlsm`/`.xlsb`/`.xlam` file in the workspace
- **Syntax highlighting** -- full TextMate grammar per the MS-VBAL specification
- **Go to Definition** (F12), **Find All References** (Shift+F12), **Rename Symbol** (F2)
- **Document outline** -- breadcrumbs and Outline panel list every Sub/Function/Property
- **Run macro** -- F5 inside a Sub runs it in Excel (Windows, requires Excel open)
- **Add / Rename / Delete** modules via right-click menus
- **Export to folder** -- right-click a workbook and choose Export All Modules
- **GitHub Copilot tools** -- Copilot can list, read, and write modules and cells from chat

---

## Configuration

| Setting | Default | Description |
|---|---|---|
| `xlide.pythonPath` | `""` | Full path to Python executable. Leave blank to use the system `python`. |
| `xlide.attachToRunningExcel` | `true` | Windows: attach to a running Excel instance when opening a workbook. |

---

## Troubleshooting

**"Python was not found"** -- Install Python and make sure `python --version` works in a terminal.
Set `xlide.pythonPath` in VS Code settings if Python is not on your PATH.

**"Required Python packages are missing"** -- Run `XLIDE: Install Python Dependencies` from the Command Palette.

**XLIDE panel is empty** -- Open a *folder* (not a single file) that contains `.xlsm`/`.xlsb`/`.xlam` files.

**Changes not saved** -- Press Ctrl+S while the module editor tab is focused.
Check the **XLIDE** output channel for error details.

---

For development docs and architecture details, see the
[GitHub repository](https://github.com/WilliamSmithEdward/xlide_vscode).
