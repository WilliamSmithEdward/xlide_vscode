import { fileURLToPath } from 'url';
import { defineConfig } from 'vitest/config';

// `npm run test:office`: the checks only real Office can answer. Like the VBE
// oracle and verify:excel:formats, a developer check and not part of CI - it
// needs Windows with Office installed, and a host it cannot use safely is
// skipped rather than failed (tests/office/officeHarness.ts says when).
//
// The suite runs the product's own script builders, coordinator and engine
// against scratch copies of the fixtures. It exists because unit tests read
// the scripts as text: a script that joined `}; else` passed every one of
// them and failed the first time a real Word ran it.
export default defineConfig({
    resolve: {
        alias: { vscode: fileURLToPath(new URL('./tests/office/vscodeStub.ts', import.meta.url)) },
    },
    test: {
        environment: 'node',
        include: ['tests/office/**/*.office.ts'],
        globalSetup: ['tests/office/officeGlobalSetup.ts'],
        // Office starts, opens, saves and quits inside a single check.
        testTimeout: 300_000,
        hookTimeout: 300_000,
        // One Office at a time: the checks share running applications.
        fileParallelism: false,
        sequence: { concurrent: false },
    },
});
