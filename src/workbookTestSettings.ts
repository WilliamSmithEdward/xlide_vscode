import {
    DEFAULT_VBA_TEST_ARTIFACT_FOLDER,
    DEFAULT_VBA_TEST_ARTIFACT_RETENTION,
} from './vbaTestArtifacts';
import {
    readWorkbookSettings,
    resolveWorkbookSetting,
    settingsPathForWorkbook,
    updateWorkbookSettings,
    type WorkbookSettingSource,
    type WorkbookSettingsConfig,
    type WorkbookTestSettingsConfig,
} from './workbookSettings';

export type WorkbookTestSettingsSource = WorkbookSettingSource;

export interface EffectiveWorkbookTestSettings {
    artifactFolder: string;
    artifactFolderSource: WorkbookTestSettingsSource;
    artifactRetention: number;
    artifactRetentionSource: WorkbookTestSettingsSource;
    settingsPath: string;
}

export interface WorkbookTestSettingsPatch {
    artifactFolder?: string;
    artifactRetention?: number;
}

export async function effectiveWorkbookTestSettings(
    workbookPath: string,
): Promise<EffectiveWorkbookTestSettings> {
    return effectiveWorkbookTestSettingsFromConfig(
        workbookPath,
        await readWorkbookSettings(workbookPath),
    );
}

export function effectiveWorkbookTestSettingsFromConfig(
    workbookPath: string,
    config: WorkbookSettingsConfig,
): EffectiveWorkbookTestSettings {
    const artifactFolder = resolveWorkbookSetting(config.tests?.artifactFolder, {
        value: DEFAULT_VBA_TEST_ARTIFACT_FOLDER,
        source: 'default',
    });
    const artifactRetention = resolveWorkbookSetting(config.tests?.artifactRetention, {
        value: DEFAULT_VBA_TEST_ARTIFACT_RETENTION,
        source: 'default',
    });
    return {
        artifactFolder: artifactFolder.value,
        artifactFolderSource: artifactFolder.source,
        artifactRetention: artifactRetention.value,
        artifactRetentionSource: artifactRetention.source,
        settingsPath: settingsPathForWorkbook(workbookPath),
    };
}

export async function updateWorkbookTestSettings(
    workbookPath: string,
    patch: WorkbookTestSettingsPatch,
): Promise<EffectiveWorkbookTestSettings> {
    await updateWorkbookSettings(workbookPath, (existing) =>
        workbookSettingsWithTestPatch(existing, patch),
    );
    return effectiveWorkbookTestSettings(workbookPath);
}

export function workbookSettingsWithTestPatch(
    existing: WorkbookSettingsConfig,
    patch: WorkbookTestSettingsPatch,
): WorkbookSettingsConfig {
    return {
        ...existing,
        tests: compactTestSettings({
            ...existing.tests,
            ...patch,
        }),
    };
}

function compactTestSettings(
    settings: WorkbookTestSettingsConfig,
): WorkbookTestSettingsConfig | undefined {
    const compacted: WorkbookTestSettingsConfig = {};
    if (settings.artifactFolder !== undefined) {
        compacted.artifactFolder = settings.artifactFolder;
    }
    if (settings.artifactRetention !== undefined) {
        compacted.artifactRetention = settings.artifactRetention;
    }
    return Object.keys(compacted).length > 0 ? compacted : undefined;
}
