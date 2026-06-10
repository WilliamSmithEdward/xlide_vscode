import {
    writeVbaTestRunArtifacts,
    type VbaTestCiStatus,
    type VbaTestRunArtifactWriteResult,
} from './vbaTestArtifacts';
import type { VbaTestHostOracleEvent } from './vbaTestHostOracle';
import type { VbaTestRunReport } from './vbaTestRunner';
import {
    effectiveWorkbookTestSettings,
    type EffectiveWorkbookTestSettings,
} from './workbookTestSettings';
import { errorMessage } from './util/errors';

export interface AgentVbaTestArtifactSettingsPayload {
    artifactFolder: string;
    artifactFolderSource: EffectiveWorkbookTestSettings['artifactFolderSource'];
    artifactRetention: number;
    artifactRetentionSource: EffectiveWorkbookTestSettings['artifactRetentionSource'];
    settingsPath: string;
}

export interface AgentVbaTestArtifactSuccessPayload {
    ok: true;
    outputFolder: string;
    runDirectory: string;
    summaryPath: string;
    hostTracePath: string;
    outputLogPath: string;
    statusPath: string;
    runId: string;
    relativePaths: VbaTestRunArtifactWriteResult['relativePaths'];
    ciStatus: VbaTestCiStatus;
    settings: AgentVbaTestArtifactSettingsPayload;
}

export interface AgentVbaTestArtifactErrorPayload {
    ok: false;
    error: string;
}

export type AgentVbaTestArtifactPayload =
    | AgentVbaTestArtifactSuccessPayload
    | AgentVbaTestArtifactErrorPayload;

export async function writeAgentVbaTestArtifacts(
    report: VbaTestRunReport,
    hostEvents: readonly VbaTestHostOracleEvent[],
): Promise<AgentVbaTestArtifactPayload> {
    try {
        const settings = await effectiveWorkbookTestSettings(report.filePath);
        const artifacts = await writeVbaTestRunArtifacts(report, hostEvents, {
            outputFolder: settings.artifactFolder,
            retention: settings.artifactRetention,
        });
        return agentVbaTestArtifactPayload(artifacts, settings);
    } catch (error) {
        return {
            ok: false,
            error: errorMessage(error),
        };
    }
}

export function agentVbaTestArtifactPayload(
    artifacts: VbaTestRunArtifactWriteResult,
    settings: EffectiveWorkbookTestSettings,
): AgentVbaTestArtifactSuccessPayload {
    return {
        ok: true,
        outputFolder: artifacts.outputFolder,
        runDirectory: artifacts.runDirectory,
        summaryPath: artifacts.summaryPath,
        hostTracePath: artifacts.hostTracePath,
        outputLogPath: artifacts.outputLogPath,
        statusPath: artifacts.statusPath,
        runId: artifacts.runId,
        relativePaths: artifacts.relativePaths,
        ciStatus: artifacts.ciStatus,
        settings: {
            artifactFolder: settings.artifactFolder,
            artifactFolderSource: settings.artifactFolderSource,
            artifactRetention: settings.artifactRetention,
            artifactRetentionSource: settings.artifactRetentionSource,
            settingsPath: settings.settingsPath,
        },
    };
}
