import type { PythonBridge } from './pythonBridge';
import {
    normalizeVbaTestSupportModuleSource,
    XLIDE_ASSERT_MODULE_NAME,
    XLIDE_ASSERT_MODULE_SOURCE,
} from './vbaTestSupportModule';

export interface VbaTestSupportStatus {
    state: 'missing' | 'blocked' | 'installed' | 'outdated' | 'unknown';
    title: string;
    description: string;
    actionLabel: string;
    canInstall: boolean;
    canRun: boolean;
}

export async function getVbaTestSupportStatus(
    bridge: PythonBridge,
    filePath: string,
): Promise<VbaTestSupportStatus> {
    try {
        const modules = await bridge.call<Array<{ name: string; type: string }>>(
            'listModules',
            { path: filePath },
        );
        const existing = modules.find(
            (module) => module.name.toLowerCase() === XLIDE_ASSERT_MODULE_NAME.toLowerCase(),
        );
        if (!existing) {
            return {
                state: 'missing',
                title: 'XlideAssert.bas Not Installed',
                description: 'The bundled test support module must be installed before XLIDE can run workbook tests.',
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
                description: 'Workbook tests can run through the XLIDE-owned read-only Excel test host.',
                actionLabel: 'Installed',
                canInstall: false,
                canRun: true,
            };
        }
        return {
            state: 'outdated',
            title: 'XlideAssert.bas Needs Update',
            description: 'The workbook has an XlideAssert standard module, but it does not match the bundled XLIDE test support module.',
            actionLabel: 'Update',
            canInstall: true,
            canRun: false,
        };
    } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        return {
            state: 'unknown',
            title: 'Test Support Unknown',
            description: `XLIDE could not inspect the workbook test support module: ${message}`,
            actionLabel: 'Refresh',
            canInstall: false,
            canRun: false,
        };
    }
}

function moduleSourceFromReadResult(result: { source?: string } | string): string {
    return typeof result === 'string' ? result : result.source ?? '';
}
