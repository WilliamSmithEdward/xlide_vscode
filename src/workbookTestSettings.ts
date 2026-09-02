import {
    DEFAULT_VBA_TEST_ARTIFACT_FOLDER,
    DEFAULT_VBA_TEST_ARTIFACT_RETENTION,
} from './vbaTestArtifacts';
import {
    readProjectSettings,
    resolveProjectSetting,
    settingsPathForProject,
    type ProjectSettingSource,
    type ProjectSettingsConfig,
} from './projectSettings';

export type WorkbookTestSettingsSource = ProjectSettingSource;

export interface EffectiveWorkbookTestSettings {
    artifactFolder: string;
    artifactFolderSource: WorkbookTestSettingsSource;
    artifactRetention: number;
    artifactRetentionSource: WorkbookTestSettingsSource;
    settingsPath: string;
}

export async function effectiveWorkbookTestSettings(
    projectPath: string,
): Promise<EffectiveWorkbookTestSettings> {
    return effectiveWorkbookTestSettingsFromConfig(
        projectPath,
        await readProjectSettings(projectPath, { lenient: true }),
    );
}

export function effectiveWorkbookTestSettingsFromConfig(
    projectPath: string,
    config: ProjectSettingsConfig,
): EffectiveWorkbookTestSettings {
    const artifactFolder = resolveProjectSetting(config.tests?.artifactFolder, {
        value: DEFAULT_VBA_TEST_ARTIFACT_FOLDER,
        source: 'default',
    });
    const artifactRetention = resolveProjectSetting(config.tests?.artifactRetention, {
        value: DEFAULT_VBA_TEST_ARTIFACT_RETENTION,
        source: 'default',
    });
    return {
        artifactFolder: artifactFolder.value,
        artifactFolderSource: artifactFolder.source,
        artifactRetention: artifactRetention.value,
        artifactRetentionSource: artifactRetention.source,
        settingsPath: settingsPathForProject(projectPath),
    };
}

