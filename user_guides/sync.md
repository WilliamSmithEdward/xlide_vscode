# XLIDE Import and Export

XLIDE can synchronize a file's VBA modules with ordinary `.bas` and `.cls`
files. Use this when you want Git history, code review, backups, or external
tooling over module text. Every macro container exports and imports:
projects, documents, presentations, and Access databases alike. A module
written into an Access database takes effect when Access next opens it and
recompiles.

For normal editing, the file and `xlide-vba` editor documents remain the
source of truth. Open modules from the XLIDE tree, edit in VS Code, and
save with Ctrl+S. Exported files are sync artifacts unless you explicitly
import them back.

Import/export is file-scoped. If your VS Code project contains multiple
macro files, choose the target in the XLIDE tree or run the command from
an open module in that file. The preview, settings sidecar, apply action,
and any writes belong to that one selected file.

## When to Use Sync

Import/export is useful for:

- putting VBA code under source control
- reviewing VBA changes in normal file diffs
- backing up all of a file's modules to a folder
- applying a reviewed batch of `.bas` or `.cls` file edits back into the file
- letting external tools inspect module text without opening the file

Import/export uses XLIDE's built-in engine and does not require Office COM.
Office COM is only required for workflows that execute VBA, such as running
macros or unit tests.

## Export Project Modules

Use **Export All Modules to Folder** when you want a folder copy of the
project's modules. XLIDE reads the live project modules, opens a preview, and
lets you inspect the file changes before applying them.

In the export preview you can:

- choose or change the export folder
- switch export mode
- inspect project-vs-file diffs
- check or uncheck individual modules
- apply only the selected changes

Standard modules export as `.bas` files. Class modules, document modules, and
UserForm code-behind modules export as `.cls` files.

Use **Export/Sync Current Module** when you only want to export the active
module instead of the whole project.

## Export Modes

Export mode is project-specific:

| Mode | UI label | Behavior |
|---|---|---|
| `exportAll` | Export All (No Deletes) | Creates missing files and updates changed files for every project module. Stale files are left alone. |
| `trueUp` | Export All + Delete Missing | Does everything `exportAll` does, then previews stale root `.bas` and `.cls` module files for removal when they no longer exist in the project. |

`trueUp` only treats root `.bas` and `.cls` files in the selected folder as
module sync files. Nested files, other file types, and `.frm` designer files are
outside import/export sync.

## Import Modules from Files

Use **Import Modules from Folder** when you want `.bas` or `.cls` files applied
back into a project. XLIDE reads the configured or selected folder and opens
the same preview experience before writing to the project.

With multiple projects in the same VS Code workspace, import still targets only
the selected project. XLIDE does not batch-import the same folder into every
detected project.

Import can:

- update existing project modules
- create missing standard modules
- create missing class modules
- skip unsafe or unsupported module creations without failing the whole import

Document modules and UserForm `.cls` code-behind modules can be updated when the
project already contains a same-named module. XLIDE cannot create missing
document modules or UserForm designer-backed modules directly from import, so
those rows show **Skipping import** in the preview and are skipped on apply.

`.frm` designer files are ignored by v2 import/export sync.

## Import Modes

Import mode is project-specific:

| Mode | UI label | Behavior |
|---|---|---|
| `updateOnly` | Import/Update (No Deletes) | Updates existing modules and creates supported missing standard/class modules. Project-only modules are left alone. |
| `trueUpStandardClass` | Import/Update + Delete Missing | Does everything `updateOnly` does, then previews project-only standard/class modules for deletion when they are missing from the folder. |

Document modules and UserForm code-behind modules are excluded from
`trueUpStandardClass` removals.

## Project Sync Settings

XLIDE stores project-specific sync settings beside the project in
`<file>.xlide_settings.json`, for example
`Budget.xlsm.xlide_settings.json`.

The sync preview is the safest place to change these values because it preserves
unrelated settings and shows whether each value came from the project sidecar,
built-in defaults, or the current session.

```json
{
  "exportFolder": "C:/absolute/path/to/export/folder",
  "exportMode": "exportAll",
  "importMode": "updateOnly"
}
```

If the sidecar contains invalid JSON, unknown keys, or invalid sync modes, XLIDE
reports the settings file as invalid instead of silently ignoring it.

## Compare Modules With Git

An Office file is one binary file to git, so the Source Control view can only
say that it changed. XLIDE reads the committed file the same way it reads the
one on disk, so you can see what changed inside it without exporting anything.

Right-click a module in the XLIDE tree:

- **Compare Module with Git HEAD** opens a diff of the module as last
  committed against the module as it is now, unsaved edits included.
- **Compare Module with Git Revision...** lists the commits that touched the
  file and diffs the module against the one you pick.

Right-click the file itself and choose **Compare File with Git HEAD** to list
every module that was modified, added, or removed since the last commit, and
open the diff of any of them.

**Show Module History** answers a question git cannot: which commits changed
this module? Git's log lists every commit that touched the file. XLIDE reads
the module out of each of the last fifty of them and keeps the commits where
its text moved, newest first, with what each did (modified, added, removed).
Pick one and the diff shows the module before that commit against the module
in it.

**Restore Module from Git HEAD** puts one module back to its committed text
and leaves the rest of the file alone. The committed text is applied as an
edit in the module's editor, so Undo brings the current text back, and the
module is saved for you unless it had unsaved edits, in which case the save is
yours. A module the last commit does not have cannot be restored; delete it if
you want the committed state.

The same three commands are in the Command Palette while a module is open.
They need `git` on your PATH, or the `git.path` setting VS Code's own git
support uses, and the file has to be committed at least once. A VB6 project's
modules are files of their own, so each one is compared with its own history.

The tree shows the same comparison without asking: a module that differs from
the last commit carries an `M` badge, one the last commit does not have an
`A`, and the file row counts them, in the colours the Explorer uses for
changed files. The marks are computed once per change to the file or the
repository, so a large file costs one parse of its committed copy, not one
per row.

## Safety Notes

Import/export is preview-first. The project or file system is changed only
after you apply the preview.

Before a large import or any delete-enabled true-up, use source control or a
project backup so the change is easy to review or roll back.

Project imports are project mutations. If a project is signed, saving module
changes can invalidate the signature; XLIDE reports signature loss the same way
it does for other project write operations.

If a write fails because the project is locked by another application, close
the other copy and retry.

Project-to-project module transfer is not a v2 import/export workflow. Treat
that as a separate explicit operation with its own source/destination selection,
conflict handling, and preview.

## Agent and Automation Notes

AI agents should normally read and write project VBA through
`xlide_readModule` and `xlide_writeModule`. Use `xlide_exportModules` only when
the task is specifically to create or refresh exported files.

The v2 agent tools include:

- `xlide_exportModules` for exporting all project modules to a folder
- `xlide_importModules` for applying a folder's `.bas`, `.cls` and `.frm`
  files to the project, under the same rules as the preview: standard and
  class modules are updated or created, document and form modules are updated
  when the project has them, and the confirmation says what would change
  before anything is written
- `xlide_configureExportMode` for setting the project-specific export mode

After a meaningful import, run project analysis and any available project tests
before treating the project as release-ready.

## Troubleshooting

- If XLIDE cannot find a folder, choose one from the preview.
- If the sidecar is invalid, fix `<file>.xlide_settings.json` or remove the
  invalid fields and reopen the preview.
- If an import row says **Skipping import**, the file represents a document or
  UserForm code-behind module that does not already exist in the project.
- If a project write fails, check project protection, open-state conflicts,
  file permissions, and the XLIDE support diagnostics.
