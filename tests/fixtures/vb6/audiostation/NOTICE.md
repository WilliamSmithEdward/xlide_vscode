# Audiostation forms (VB6 fixture)

Source: https://github.com/Sibra-Soft/audiostation, branch `master`,
commit `8dcf55f3cb9783cb17224bee6853a8ab50bfc5c0`, folder `Source/`.

License: MIT, Copyright (c) 2009 Sibra-Soft. The full text is in `LICENSE`
beside this file; it was read in full before these files were copied.

Files, unchanged:

- `Form_OpenDialog.frm`/`.frx`: OCX references (`Object =` lines), an SSTab
  with indexed and dotted property keys (`Tab(0).Control(0).Enabled`), an
  ImageList whose images are `BeginProperty` groups carrying class ids, and a
  `TabPicture(0)` sidecar reference.
- `Form_Settings_Record.frm`/`.frx`: ComboBox rows (`List`, `ItemData`).
- `Form_Track_Properties.frm`/`.frx`: a short-string `Text` sidecar record.
- `Hyperlink.ctl`/`.ctx`: a UserControl with picture-valued properties in its
  own sidecar.
- `MixSlider.ctl`/`.ctx`: a UserControl hosting other controls, whose headers
  carry extender properties (`Object.Width`, `Object.Visible`).

These files use LF line endings as published.

These files are test fixtures only. They are excluded from the extension
package by `.vscodeignore` (`tests/**`).
