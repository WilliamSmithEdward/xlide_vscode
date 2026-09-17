import type { ProjectEngine } from './projectEngine';
import {
    normalizeVbaTestSupportModuleSource,
    XLIDE_ASSERT_MODULE_NAME,
    XLIDE_ASSERT_MODULE_SOURCE,
} from './vbaTestSupportModule';
import { errorMessage } from './util/errors';
import { containerAppNameForPath } from './macroContainerUi';

export interface VbaTestSupportStatus {
    state: 'missing' | 'blocked' | 'installed' | 'outdated' | 'unknown';
    title: string;
    description: string;
    actionLabel: string;
    canInstall: boolean;
    canRun: boolean;
}

/** The project's module named like the bundled test support module, whatever its kind. */
export async function findVbaTestSupportModule(
    bridge: ProjectEngine,
    filePath: string,
): Promise<{ name: string; type: string } | undefined> {
    const modules = await bridge.call<Array<{ name: string; type: string }>>(
        'listModules',
        { path: filePath },
    );
    return modules.find(
        (module) => module.name.toLowerCase() === XLIDE_ASSERT_MODULE_NAME.toLowerCase(),
    );
}

export async function getVbaTestSupportStatus(
    bridge: ProjectEngine,
    filePath: string,
): Promise<VbaTestSupportStatus> {
    try {
        const existing = await findVbaTestSupportModule(bridge, filePath);
        if (!existing) {
            return {
                state: 'missing',
                title: 'XlideAssert.bas Not Installed',
                description: 'The bundled test support module must be installed before XLIDE can run tests in this file.',
                actionLabel: 'Install',
                canInstall: true,
                canRun: false,
            };
        }
        if (existing.type !== 'standard') {
            return {
                state: 'blocked',
                title: `${XLIDE_ASSERT_MODULE_NAME} Name Conflict`,
                description: `"${XLIDE_ASSERT_MODULE_NAME}" exists as a ${existing.type} module. Rename it before installing the XLIDE test support module.`,
                actionLabel: 'Blocked',
                canInstall: false,
                canRun: false,
            };
        }

        const current = await bridge.call<{ source?: string } | string>(
            'readModule',
            { path: filePath, module: existing.name },
        );
        const installed = normalizeVbaTestSupportModuleSource(moduleSourceFromReadResult(current)) ===
            normalizeVbaTestSupportModuleSource(XLIDE_ASSERT_MODULE_SOURCE);
        if (installed) {
            return {
                state: 'installed',
                title: 'XlideAssert.bas Installed',
                description: `Tests can run through the XLIDE-owned read-only ${containerAppNameForPath(filePath)} test host.`,
                actionLabel: 'Installed',
                canInstall: false,
                canRun: true,
            };
        }
        return {
            state: 'outdated',
            title: 'XlideAssert.bas Needs Update',
            description: 'The file has an XlideAssert standard module, but it does not match the bundled XLIDE test support module.',
            actionLabel: 'Update',
            canInstall: true,
            canRun: false,
        };
    } catch (err) {
        const message = errorMessage(err);
        return {
            state: 'unknown',
            title: 'Test Support Unknown',
            description: `XLIDE could not inspect the test support module: ${message}`,
            actionLabel: 'Refresh',
            canInstall: false,
            canRun: false,
        };
    }
}

function moduleSourceFromReadResult(result: { source?: string } | string): string {
    return typeof result === 'string' ? result : result.source ?? '';
}
