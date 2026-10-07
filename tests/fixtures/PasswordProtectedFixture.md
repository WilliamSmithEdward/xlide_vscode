# Password-protected Office fixtures

The public test password for every file below is **Test66**.

- `PasswordProtectedFixture.xlsm`: unchanged native locked workbook from `xlide_vbide/tools/harness/fixtures/LockedProjectFixture.xlsm`, commit `a1091496bdfa615904af89922812150792f57947`. Contains `Runner`, `Helper`, and the workbook's document modules. Its existing Office password verifies in XLIDE.
- `PasswordProtectedFormFixture.xlsm`: the Office-authored `FormFixture.xlsm`, retaining its `FrmPicker` designer and code, with the native workbook's exact `ID`, `CMG`, `DPB`, and `GC` protection records transferred into the PROJECT header. No code or designer is regenerated.
- `PasswordProtectedWordFixture.docm`, `PasswordProtectedPowerPointFixture.pptm`, and `PasswordProtectedAccessFixture.accdb`: copies of the repository's corresponding Office-authored fixtures, with MS-OVBA password records added by the test-only helper. The viewing lock uses the native fixture's `fVBEProtected` flag and hidden-project visibility state. Access records are written into its actual named storage row.

Open copies for manual testing. Cancelled or incorrect passwords must deny XLIDE access; `Test66` must unlock the selected file once for the current VS Code session. Close Office copies without saving to retain their initial protected state. These fixtures are not production documents.

Regression tests: `npx --no-install vitest run tests/projectProtection.test.ts`.

Real VS Code integration: `npm run compile`, then `npx --no-install vscode-test --config tests/projectPasswordIntegration.config.mjs`. The runner pins VS Code 1.139.1; `XLIDE_TEST_CODE_EXECUTABLE` can select an existing compatible installation. The suite uses the activated extension's filesystem, editor saves, UserForm designer, commands and agent tools; user-dialog answers are stubbed.

Live Office checks use an automation-created baseline and report the actual VBE add-in connection state. Process-launched add-in coexistence needs a separate manual check. Installation alone does not prove XLIDE VBIDE loaded in an automation-created host. Check the actual add-in connection state when testing coexistence.
