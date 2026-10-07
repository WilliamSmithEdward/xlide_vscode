# Protected VBA fixture

`binaries/PasswordProtectedFixture.xlsm` is a permanent, intentionally password-protected Excel workbook for issue #1298. Its public test password is **XLIDE-test-1298!**. It contains `ProtectedProbe` (a harmless function returning 1298), `ProtectedForm` (an empty UserForm), and the blank template's document modules. It has no auto-run macro.

The fixture starts from the repository's Office-authored `assets/templates/blank.xlsm`. `tests/helpers/projectProtectionFixture.ts` adds the probe/form through XLIDE and writes deterministic MS-OVBA hashed-password, viewing-lock, and visibility records, with the required zero project CLSID. It is synthetic test protection, not a claim of a new Office-authored binary. Tests copy it before making changes; do not save over the committed fixture.

Regression tests: `npx --no-install vitest run tests/projectProtection.test.ts`.

Real VS Code integration: `npm run compile`, then `npx --no-install vscode-test --config tests/projectPasswordIntegration.config.mjs`. The runner pins VS Code 1.139.1. `XLIDE_TEST_CODE_EXECUTABLE` can point to an existing compatible installation. The suite uses the activated extension's virtual filesystem, editor saves, UserForm designer, and agent tools; only the user input dialog is stubbed, following the existing integration-test pattern.
