import type { PythonBridge } from './pythonBridge';
import { checkExcelComAvailability, type ExcelComAvailabilityStatus } from './excelComAvailability';
import { getVbaTestSupportStatus, type VbaTestSupportStatus } from './vbaTestSupportStatus';
import {
    runWorkbookVbaTests,
    type VbaTestProgressReporter,
    type VbaTestRunExecution,
    type VbaTestRunOptions,
} from './vbaTestExecution';
import {
    writeVbaTestRunArtifacts,
    type VbaTestRunArtifactWriteResult,
} from './vbaTestArtifacts';
import {
    effectiveWorkbookTestSettings,
    type EffectiveWorkbookTestSettings,
} from './workbookTestSettings';
import { errorMessage } from './util/errors';

export type VbaTestRunPipelineArtifacts =
    | { ok: true; artifacts: VbaTestRunArtifactWriteResult; settings: EffectiveWorkbookTestSettings }
    | { ok: false; error: string };

export type VbaTestRunPipelineResult =
    | { kind: 'blocked-support'; support: VbaTestSupportStatus }
    | { kind: 'blocked-com'; runtime: ExcelComAvailabilityStatus }
    | { kind: 'completed'; execution: VbaTestRunExecution; artifacts: VbaTestRunPipelineArtifacts };

export type VbaTestRunPipelineRunner = (progress?: VbaTestProgressReporter) => Promise<VbaTestRunExecution>;

export interface ExecuteVbaTestRunOptions extends VbaTestRunOptions {
    log?: (message: string) => void;
    // Lets a caller wrap the host run (e.g. in VS Code progress UI); the run
    // executes directly when omitted.
    runTests?: (run: VbaTestRunPipelineRunner) => PromiseLike<VbaTestRunExecution>;
}

// Single coordinator for the VBA test run pipeline (support gate, Excel COM
// gate, host run, artifact write) shared by the command-palette and agent
// tool entry points. Callers adapt the result to their own presentation.
export async function executeVbaTestRun(
    bridge: PythonBridge,
    filePath: string,
    options: ExecuteVbaTestRunOptions = {},
): Promise<VbaTestRunPipelineResult> {
    const support = await getVbaTestSupportStatus(bridge, filePath);
    if (!support.canRun) {
        return { kind: 'blocked-support', support };
    }
    const runtime = await checkExcelComAvailability();
    if (!runtime.canRun) {
        return { kind: 'blocked-com', runtime };
    }
    const runTests = options.runTests ?? ((run) => run());
    const execution = await runTests((progress) => runWorkbookVbaTests(bridge, filePath, {
        selection: options.selection,
        failFast: options.failFast,
        log: options.log,
        progress,
    }));
    const artifacts = await writeVbaTestRunPipelineArtifacts(execution);
    return { kind: 'completed', execution, artifacts };
}

export async function writeVbaTestRunPipelineArtifacts(
    execution: VbaTestRunExecution,
): Promise<VbaTestRunPipelineArtifacts> {
    try {
        const settings = await effectiveWorkbookTestSettings(execution.report.filePath);
        const artifacts = await writeVbaTestRunArtifacts(execution.report, execution.hostEvents, {
            outputFolder: settings.artifactFolder,
            retention: settings.artifactRetention,
        });
        return { ok: true, artifacts, settings };
    } catch (err) {
        return { ok: false, error: errorMessage(err) };
    }
}
