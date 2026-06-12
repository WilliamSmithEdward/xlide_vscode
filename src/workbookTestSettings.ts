import {
    DEFAULT_VBA_TEST_ARTIFACT_FOLDER,
    DEFAULT_VBA_TEST_ARTIFACT_RETENTION,
} from './vbaTestArtifacts';
import {
    readWorkbookSettings,
    resolveWorkbookSetting,
    settingsPathForWorkbook,
    type WorkbookSettingSource,
    type WorkbookSettingsConfig,
} from './workbookSettings';

export type WorkbookTestSettingsSource = WorkbookSettingSource;

export interface EffectiveWorkbookTestSettings {
    artifactFolder: string;
    artifactFolderSource: WorkbookTestSettingsSource;
    artifactRetention: number;
    artifactRetentionSource: WorkbookTestSettingsSource;
    settingsPath: string;
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

