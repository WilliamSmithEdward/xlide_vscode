import type { VbaTestCase } from './vbaTestRunner';
import { XLIDE_TEST_HOST_EVENT_PREFIX } from './vbaTestHostOracle';
import { readExtensionTextAsset } from './extensionAssets';
import { XLIDE_TEST_RUNNER_MODULE_NAME } from './vbaTestRunnerModuleCodegen';
import { psSingleQuoted } from './util/powershell';

export const DEFAULT_VBA_TEST_TIMEOUT_MS = 30000;

export interface VbaTestHostPlanItem {
    qualifiedName: string;
    timeoutMs: number;
    expectedFailure: boolean;
}

export interface OwnedReadOnlyExcelTestHostScriptOptions {
    failFast?: boolean;
    runnerModuleName?: string;
}

export function vbaTestHostPlanItems(tests: readonly VbaTestCase[]): VbaTestHostPlanItem[] {
    return tests.map((test) => ({
        qualifiedName: test.qualifiedName,
        timeoutMs: test.metadata.timeoutMs ?? DEFAULT_VBA_TEST_TIMEOUT_MS,
        expectedFailure: Boolean(test.metadata.xfailReason),
    }));
}

// The C# modal watcher is a pure static asset: it ships as
// assets/testhost/XlideTestModalWatcher.cs and is read from the installed
// extension layout at runtime (see extensionAssets.ts).
function productionModalWatcherCSharp(): string {
    return readExtensionTextAsset('assets/testhost/XlideTestModalWatcher.cs');
}

export function buildOwnedReadOnlyExcelTestHostScript(
    filePath: string,
    tests: readonly VbaTestHostPlanItem[],
    options: OwnedReadOnlyExcelTestHostScriptOptions = {},
): string {
    const testsJson = JSON.stringify(tests);
    const runnerModuleName = options.runnerModuleName ?? XLIDE_TEST_RUNNER_MODULE_NAME;
    const modalWatcherSource = productionModalWatcherCSharp();
    // Dynamic preamble + static body (assets/testhost/run-vba-tests.ps1),
    // joined with newlines so PowerShell error positions point at meaningful
    // lines of the staged run-vba-tests.ps1.
    return [
        '$ErrorActionPreference = "Stop"',
        '$ProgressPreference = "SilentlyContinue"',
        `$targetPath = ${psSingleQuoted(filePath)}`,
        `$testsJson = ${psSingleQuoted(testsJson)}`,
        `$runnerModuleName = ${psSingleQuoted(runnerModuleName)}`,
        `$failFast = ${options.failFast ? '$true' : '$false'}`,
        `$eventPrefix = ${psSingleQuoted(XLIDE_TEST_HOST_EVENT_PREFIX)}`,
        `$modalWatcherSource = ${psSingleQuoted(modalWatcherSource)}`,
        readExtensionTextAsset('assets/testhost/run-vba-tests.ps1'),
    ].join('\n');
}
