# Third-party notices

XLIDE's VB6 host object model (`src/analyzer/host/vb6ObjectModelData.ts`,
bundled into the extension) carries member descriptions transcribed from the
following sources. How the transcription works is described in
`docs/vb6_reference_data.md`.

## twinBASIC documentation

The `VB` library metadata (App, Clipboard, Forms, Global, Printer, Printers,
Screen, Form, MDIForm, Menu, PropertyPage, UserControl and the intrinsic
controls) is transcribed from the package documentation in
https://github.com/twinbasic/documentation, commit
`5799ad8e236a77d2b9431f56ab3346fabcbefc12` (2026-07-17), files
`docs/Reference/Default/VB/<Class>/index.md`. Member descriptions are the
pages' own sentences, plain-texted. That repository's license, read in full
before the transcription, is reproduced here:

    MIT License

    Copyright (c) 2022 twinbasic

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

## Microsoft Visual Basic 6.0 documentation

The transcription is cross-read against the names of the pages in the table
of contents of Microsoft's archived Visual Basic 6.0 documentation
(https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/),
fetched on 2026-09-02. Only page names are used; no page text is reproduced.

## msvbvm60.dll

The `VBRUN` library metadata is read from the type library inside the Visual
Basic 6.0 runtime that Windows ships (`msvbvm60.dll`, file version
6.0.98.48): type, member and constant names, signatures, values and the help
strings the library carries. The runtime itself is not bundled.
