// Integration tests: a real VS Code, this extension in development mode, and
// a throwaway workspace holding a copy of the fixture workbook. Run with
// `npm run test:integration`. The first run downloads a VS Code build into
// .vscode-test/, which takes minutes rather than seconds.
import { defineConfig } from '@vscode/test-cli';
import { cpSync, existsSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { fileURLToPath, URL } from 'node:url';

const cache = new URL('./.vscode-test/', import.meta.url);
mkdirSync(cache, { recursive: true });

/** A folder holding a copy of the fixture workbook and a loose module. */
function seedWorkspace(folder) {
  mkdirSync(folder, { recursive: true });
  cpSync(
    new URL('./tests/fixtures/binaries/FormFixture.xlsm', import.meta.url),
    new URL('FormFixture.xlsm', folder),
  );
  writeFileSync(new URL('Loose.bas', folder), [
    'Attribute VB_Name = "Loose"',
    'option explicit',
    'sub hello()',
    'dim x as long',
    'if x=1 then',
    'msgbox "hi"',
    'end if',
    'end sub',
    '',
  ].join('\r\n'));
}

// Every run starts from the same workbook: the suites write modules into it,
// commit it and change it again, so a leftover from the last run would make
// the assertions depend on history.
const workspace = new URL('integration-workspace/', cache);
rmSync(workspace, { recursive: true, force: true });
seedWorkspace(workspace);

// The folder for an Extension Development Host opened by hand. It is seeded
// once and never reset: a host left open holds handles inside it (its git
// extension watches the repository), and deleting it under the host fails the
// whole run with EPERM.
const handWorkspace = new URL('workspace/', cache);
if (!existsSync(handWorkspace)) {
  seedWorkspace(handWorkspace);
}

export default defineConfig({
  label: 'integration',
  files: 'out/test/**/*.test.js',
  workspaceFolder: fileURLToPath(workspace),
  version: 'stable',
  mocha: {
    ui: 'tdd',
    timeout: 60000,
  },
  launchArgs: [
    // Other extensions would change which providers answer, making results
    // depend on whatever the developer happens to have installed.
    '--disable-extensions',
    '--user-data-dir', fileURLToPath(new URL('user-data/', cache)),
  ],
});
