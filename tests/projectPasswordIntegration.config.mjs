import { defineConfig } from '@vscode/test-cli';
import { mkdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
const workspace = new URL('../.vscode-test/password-workspace/', import.meta.url);
mkdirSync(workspace, { recursive: true });
export default defineConfig({
    label: 'project-password',
    files: '../out/test/integration.test.js',
    extensionDevelopmentPath: '..',
    workspaceFolder: fileURLToPath(workspace),
    version: '1.139.1',
    ...(process.env.XLIDE_TEST_CODE_EXECUTABLE ? { useInstallation: { fromPath: process.env.XLIDE_TEST_CODE_EXECUTABLE } } : {}),
    mocha: { ui: 'tdd', timeout: 60000, grep: 'Protected VBA project integration' },
    launchArgs: [
        '--disable-extensions', '--skip-welcome', '--skip-release-notes',
        '--user-data-dir', fileURLToPath(new URL('../.vscode-test/password-user-data/', import.meta.url)),
    ],
});
