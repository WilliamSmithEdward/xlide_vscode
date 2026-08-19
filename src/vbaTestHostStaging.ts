import * as fs from 'fs';
import * as path from 'path';
import type { WorkbookEngine } from './workbookEngine';
import type { VbaTestCase } from './vbaTestRunner';
import {
    buildOwnedReadOnlyExcelTestHostScript,
    vbaTestHostPlanItems,
    type VbaTestHostApp,
} from './vbaTestExcelHost';
import {
    buildVbaTestDirectRunnerModule,
    buildVbaTestDispatchModule,
    XLIDE_TEST_DISPATCH_MODULE_NAME,
} from './vbaTestRunnerModuleCodegen';
import { createVbaTestHostTempDir } from './vbaTestTempFiles';
import {
    XLIDE_ASSERT_MODULE_NAME,
    XLIDE_ASSERT_MODULE_SOURCE,
} from './vbaTestSupportModule';
import { errorMessage } from './util/errors';

export interface VbaTestHostStagingOptions {
    failFast?: boolean;
    /** Which Office application hosts the run. Defaults to Excel. */
    hostApp?: VbaTestHostApp;
    log: (message: string) => void;
}

export interface VbaTestHostStaging {
    tempWorkbookPath: string;
    hostScriptPath: string;
    /** Best-effort async removal of the staging dir; idempotent, retries once. */
    dispose(): void;
}

/**
 * Stages the owned read-only Excel test host run in a private temp dir: copies
 * the workbook, injects the XlideAssert support module and a uniquely named
 * direct-call dispatcher module via the workbook engine, and writes the
 * generated run-vba-tests.ps1 host script. On staging failure the temp dir is
 * removed before the error propagates.
 */
export async function stageOwnedReadOnlyExcelTestHost(
    bridge: WorkbookEngine,
    filePath: string,
    tests: readonly VbaTestCase[],
    options: VbaTestHostStagingOptions,
): Promise<VbaTestHostStaging> {
    const hostScriptDir = await createVbaTestHostTempDir();
    const tempWorkbookPath = path.join(hostScriptDir, path.basename(filePath));
    const hostScriptPath = path.join(hostScriptDir, 'run-vba-tests.ps1');
    const runnerModuleName = `XlideRun${Date.now().toString(36).slice(-8)}`;
    try {
        await fs.promises.copyFile(filePath, tempWorkbookPath);
        await bridge.call<{ ok?: boolean; signatureDropped?: boolean }>('writeModule', {
            path: tempWorkbookPath,
            module: XLIDE_ASSERT_MODULE_NAME,
            source: XLIDE_ASSERT_MODULE_SOURCE,
            kind: 'standard',
        });
        await bridge.call<{ ok?: boolean; signatureDropped?: boolean }>('writeModule', {
            path: tempWorkbookPath,
            module: runnerModuleName,
            source: buildVbaTestDirectRunnerModule(tests, runnerModuleName),
            kind: 'standard',
        });
        // The by-name dispatcher lets XlideAssert.Throws/DoesNotThrow run
        // their targets as direct calls, which behaves identically on every
        // host (Word never propagates a Run-target's error to the caller).
        const stagedModules = await bridge.call<Array<{ name: string; type?: string; source?: string }>>(
            'readModules',
            { path: tempWorkbookPath, full: true },
        );
        await bridge.call<{ ok?: boolean; signatureDropped?: boolean }>('writeModule', {
            path: tempWorkbookPath,
            module: XLIDE_TEST_DISPATCH_MODULE_NAME,
            source: buildVbaTestDispatchModule(stagedModules),
            kind: 'standard',
        });
        const script = buildOwnedReadOnlyExcelTestHostScript(
            tempWorkbookPath,
            vbaTestHostPlanItems(tests),
            { failFast: options.failFast, runnerModuleName, hostApp: options.hostApp },
        );
        await fs.promises.writeFile(hostScriptPath, script, 'utf8');
    } catch (err) {
        try {
            await fs.promises.rm(hostScriptDir, { recursive: true, force: true });
        } catch {
            // Best-effort cleanup only.
        }
        throw err;
    }

    let disposed = false;
    const dispose = () => {
        if (disposed) {
            return;
        }
        disposed = true;
        void fs.promises.rm(hostScriptDir, { recursive: true, force: true }).catch((err) => {
            const message = errorMessage(err);
            options.log(`[runVbaTests] Could not delete temporary host script: ${message}`);
            setTimeout(() => {
                void fs.promises.rm(hostScriptDir, { recursive: true, force: true }).catch(() => {
                    // Best-effort cleanup only.
                });
            }, 1000);
        });
    };

    return { tempWorkbookPath, hostScriptPath, dispose };
}
