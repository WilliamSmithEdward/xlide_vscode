import * as fs from 'fs';
import * as path from 'path';
import type { PythonBridge } from './pythonBridge';
import type { VbaTestCase } from './vbaTestRunner';
import {
    buildOwnedReadOnlyExcelTestHostScript,
    vbaTestHostPlanItems,
} from './vbaTestExcelHost';
import { buildVbaTestDirectRunnerModule } from './vbaTestRunnerModuleCodegen';
import { createVbaTestHostTempDir } from './vbaTestTempFiles';
import {
    XLIDE_ASSERT_MODULE_NAME,
    XLIDE_ASSERT_MODULE_SOURCE,
} from './vbaTestSupportModule';
import { errorMessage } from './util/errors';

export interface VbaTestHostStagingOptions {
    failFast?: boolean;
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
 * direct-call dispatcher module via the Python bridge, and writes the
 * generated run-vba-tests.ps1 host script. On staging failure the temp dir is
 * removed before the error propagates.
 */
export async function stageOwnedReadOnlyExcelTestHost(
    bridge: PythonBridge,
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
        const script = buildOwnedReadOnlyExcelTestHostScript(
            tempWorkbookPath,
            vbaTestHostPlanItems(tests),
            { failFast: options.failFast, runnerModuleName },
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
