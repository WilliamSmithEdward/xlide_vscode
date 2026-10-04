import { defineConfig } from '@vscode/test-cli';
import { mkdirSync } from 'node:fs';
import { createServer } from 'node:net';
import { fileURLToPath } from 'node:url';

// A debugger port on this disposable instance lets an independent process
// test renderer keys while the extension host is intentionally occupied.
const server = createServer();
await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', resolve);
});
const port = server.address().port;
await new Promise((resolve, reject) => server.close(error => error ? reject(error) : resolve()));
process.env.XLIDE_UI_DEBUG_PORT = String(port);
const workspace = new URL('../.vscode-test/native-enter-routing-workspace/', import.meta.url);
mkdirSync(workspace, { recursive: true });
export default defineConfig({
    label: 'native-enter',
    files: '../out/test/integration.test.js',
    extensionDevelopmentPath: '..',
    workspaceFolder: fileURLToPath(workspace),
    version: '1.139.1',
    ...(process.env.XLIDE_TEST_CODE_EXECUTABLE ? { useInstallation: { fromPath: process.env.XLIDE_TEST_CODE_EXECUTABLE } } : {}),
    mocha: { ui: 'tdd', timeout: 120000, grep: 'Native Enter caret ownership' },
    launchArgs: [
        '--disable-extensions', '--skip-welcome', '--skip-release-notes',
        `--remote-debugging-port=${port}`, '--remote-debugging-address=127.0.0.1',
        '--user-data-dir', fileURLToPath(new URL('../.vscode-test/native-enter-routing-user-data/', import.meta.url)),
    ],
});
