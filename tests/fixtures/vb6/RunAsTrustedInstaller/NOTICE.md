# RunAsTrustedInstaller (VB6 fixture)

Source: https://github.com/fafalone/RunAsTrustedInstaller, branch `main`,
commit `6e500b58d1c4ef62d31d23437ef0c27ece6ed3e1`, folder `VB6/`.

License: MIT, Copyright (c) 2022 fafalone. The full text is in `LICENSE`
beside this file; it was read in full before these files were copied.

Files: `Project1.vbp`, `Form1.frm`, `Form1.frx`, `modRunAsTI.bas`,
`Project1.res`, unchanged. A one-form, one-module project authored in Visual
Basic 6 itself: the ground truth for the manifest shape, a form header with
`BeginProperty` blocks, and a `.frx` holding a length-prefixed string.

These files are test fixtures only. They are excluded from the extension
package by `.vscodeignore` (`tests/**`).
