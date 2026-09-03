# Visual Basic 6 projects

XLIDE opens classic Visual Basic 6 projects beside your Office files. A `.vbp`
appears in the same tree, its files are edited as the plain text they already
are, and its forms open in the same designer the UserForms use.

Visual Basic 6 itself is not required for any of this. XLIDE reads and writes
the project, its forms, and their `.frx` sidecars on its own.

## Opening a project

Put a folder containing a `.vbp` in your workspace. XLIDE finds it the same way
it finds a workbook, and shows it in the XLIDE tree with a project icon.

Expand it and you get the modules the manifest names, each with the kind the
manifest gives it: standard modules, class modules, forms, MDI forms,
UserControls, PropertyPages and ActiveX Designers. Expand a module for its
procedures. Click one to open the file at that procedure.

The file is the module. There is no virtual copy and no import step: opening a
module opens the `.bas`, `.cls`, `.frm`, `.ctl` or `.pag` file itself, and
saving it saves that file.

## Language services

A module that a discovered project claims is analyzed as part of that project,
not as a loose file. The other modules' procedures and types are in scope, the
module's name comes from its own `VB_Name` attribute, and its kind comes from
the manifest.

Completion, hover and signature help resolve against a VB6 object model rather
than an Office one. `App`, `Screen`, `Printer`, `Clipboard` and `Forms` are
there, along with the intrinsic controls and the roughly 590 VB constants such
as `vbKeyReturn`. Inside a form, `Me` is the class the designer makes it, so a
form reaches `Show` and an MDI form reaches `Arrange`.

A form's controls are members the designer declared, not the text, so code that
touches `Command1` or `Text1` is understood. A control array's elements carry
the control's type, so `Label1(2).Caption` resolves.

The model's provenance is recorded rather than assumed. VB6's own type library
is not redistributable and is not present, so the model is built from the
Visual Basic runtime's type libraries plus a transcription of the documented
`VB` library, with everything excluded listed in the repository's reference
notes. It is treated as reported evidence: it offers and describes, and it does
not raise an error on its own.

## The form designer

Open a form's Designer row in the tree, or use Open With on the file. You get
the canvas, the properties pane, the toolbox and the form's header text, all in
one tab.

Every intrinsic control draws as itself: labels, text boxes, combo and list
boxes, check boxes, option buttons, command buttons, frames, picture boxes,
images, scroll bars, timers, lines, shapes, the drive, directory and file
lists, the Data control and the OLE container. Menus draw as a menu bar. A
control from an OCX draws at its true bounds under its own name, because XLIDE
models the intrinsic set and does not guess at a third-party control's look.

Gestures are the ones you expect: drag to move, drag a handle to resize, drop
from the toolbox to add, Delete to remove, drag into a frame to reparent, and
commands for z-order and tab order. Copy and paste duplicate a control with a
fresh name, and a container brings its children.

The properties pane lists what the header states plus the properties the VB6
designer writes for that kind. Enum properties offer VB6's own constants,
colors get a swatch and a picker, and a font group is edited field by field.
Renaming a control renames every element of its array with it, and setting
`Index` turns a control into an array element.

Double-click a control to open its default event handler in the code below,
creating the stub if it is not there. The event is one the control actually
raises: a scroll bar opens `Change` rather than `Click`, a Data control opens
`Reposition`, and a Line or Shape is refused, because VB6 raises nothing on
either.

### What the designer writes

The form's file is the document. A gesture rewrites the designer header at the
top of the file and never touches the code below it, so:

- Ctrl+Z is ordinary text undo.
- The tab carries the dirty dot until you save.
- A form the designer never changed saves byte for byte identical.
- A real change shows as a diff inside the header block and nowhere else.

The `.frx` sidecar is read for the strings, lists and pictures a header keeps
there. It is written for one thing only: a text value containing a line break,
which the header format cannot hold. That record is appended when you save the
form, so the file and its sidecar move together, and undoing the edit before
saving writes nothing. Pictures and list rows are read but never written, and a
gesture that would set one is refused rather than guessed at.

## F5

F5 in the designer, or on a VB6 module, saves every file the project claims and
then hands the `.vbp` to whatever application Windows has registered for it,
which is Visual Basic 6 where it is installed. This is the same thing F5
already does for a Word or PowerPoint project.

Where nothing is registered, F5 says so and names what Windows offers under
Open With rather than failing with a bare shell error.

XLIDE does not build or run VB6 projects. There is no compiler in the
extension and none is invoked.

## Limits worth knowing

- **A `.vbp` is recognized by its name.** A file with that extension is treated
  as a VB6 project; if it does not parse as a manifest, the failure appears
  when you expand it.
- **Adding, renaming and deleting modules is refused.** Those change the
  manifest, and XLIDE edits existing modules only. Add the file and the
  manifest entry yourself, and the tree picks it up.
- **A form outside any project in the workspace** is analyzed as a lone module,
  the same fallback a stray `.bas` gets.
- **MDI forms draw as plain forms.** The MDI relationship is recognized but not
  drawn.
- **Third-party controls are drawn as boxes**, at their real bounds and under
  their real names. Their properties still appear in the pane, from the header.
- **Sidecar records only ever append.** Setting a multi-line value repeatedly
  grows the `.frx`, because XLIDE never rewrites the whole file the way VB6
  does on its own save. Nothing breaks; the file grows.
