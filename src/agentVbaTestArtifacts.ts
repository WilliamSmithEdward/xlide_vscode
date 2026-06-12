import {
    type VbaTestCiStatus,
    type VbaTestRunArtifactWriteResult,
} from './vbaTestArtifacts';
import type { VbaTestRunPipelineArtifacts } from './vbaTestRunPipeline';
import type { EffectiveWorkbookTestSettings } from './workbookTestSettings';

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

export function agentVbaTestArtifactPayloadFromPipeline(
    artifacts: VbaTestRunPipelineArtifacts,
): AgentVbaTestArtifactPayload {
    if (!artifacts.ok) {
        return {
            ok: false,
            error: artifacts.error,
        };
    }
    return agentVbaTestArtifactPayload(artifacts.artifacts, artifacts.settings);
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
