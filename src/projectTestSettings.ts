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

export type ProjectTestSettingsSource = ProjectSettingSource;

export interface EffectiveProjectTestSettings {
    artifactFolder: string;
    artifactFolderSource: ProjectTestSettingsSource;
    artifactRetention: number;
    artifactRetentionSource: ProjectTestSettingsSource;
    settingsPath: string;
}

export async function effectiveProjectTestSettings(
    projectPath: string,
): Promise<EffectiveProjectTestSettings> {
    return effectiveProjectTestSettingsFromConfig(
        projectPath,
        await readProjectSettings(projectPath, { lenient: true }),
    );
}

export function effectiveProjectTestSettingsFromConfig(
    projectPath: string,
    config: ProjectSettingsConfig,
): EffectiveProjectTestSettings {
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

