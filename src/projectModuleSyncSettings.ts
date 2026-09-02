import type { ImportMode, ModuleSyncModeSource } from './moduleSyncPlan';
import {
    normalizeExportMode,
    normalizeImportMode,
    readProjectSettings,
    resolveProjectSetting,
    settingsPathForProject,
    updateProjectSettings,
    type ExportMode,
    type ProjectSettingsConfig,
} from './projectSettings';

export const DEFAULT_WORKBOOK_EXPORT_MODE: ExportMode = 'exportAll';
export const DEFAULT_WORKBOOK_IMPORT_MODE: ImportMode = 'updateOnly';

export type ProjectModuleSyncFolderSource = 'project' | 'session' | 'missing';
export type ProjectModuleSyncModeSource = ModuleSyncModeSource;

export interface EffectiveProjectModuleSyncSettings {
    folderPath?: string;
    folderPathSource: ProjectModuleSyncFolderSource;
    exportMode: ExportMode;
    exportModeSource: ProjectModuleSyncModeSource;
    importMode: ImportMode;
    importModeSource: ProjectModuleSyncModeSource;
    settingsPath: string;
}

export interface ProjectModuleSyncSettingsPatch {
    folderPath?: string;
    exportMode?: ExportMode;
    importMode?: ImportMode;
}

export async function effectiveProjectModuleSyncSettings(
    projectPath: string,
): Promise<EffectiveProjectModuleSyncSettings> {
    return effectiveProjectModuleSyncSettingsFromConfig(
        projectPath,
        await readProjectSettings(projectPath, { lenient: true }),
    );
}

export function effectiveProjectModuleSyncSettingsFromConfig(
    projectPath: string,
    config: ProjectSettingsConfig,
): EffectiveProjectModuleSyncSettings {
    const exportMode = resolveProjectSetting(config.exportMode, {
        value: DEFAULT_WORKBOOK_EXPORT_MODE,
        source: 'default',
    });
    const importMode = resolveProjectSetting(config.importMode, {
        value: DEFAULT_WORKBOOK_IMPORT_MODE,
        source: 'default',
    });
    const exportFolder = config.exportFolder && config.exportFolder.trim() !== ''
        ? config.exportFolder
        : undefined;
    return {
        folderPath: exportFolder,
        folderPathSource: exportFolder ? 'project' : 'missing',
        exportMode: exportMode.value,
        exportModeSource: exportMode.source,
        importMode: importMode.value,
        importModeSource: importMode.source,
        settingsPath: settingsPathForProject(projectPath),
    };
}

export async function updateProjectModuleSyncSettings(
    projectPath: string,
    patch: ProjectModuleSyncSettingsPatch,
): Promise<EffectiveProjectModuleSyncSettings> {
    const updated = await updateProjectSettings(projectPath, (existing) =>
        projectSettingsWithModuleSyncPatch(existing, patch),
    );
    return effectiveProjectModuleSyncSettingsFromConfig(projectPath, updated);
}

export async function setProjectModuleSyncExportMode(
    projectPath: string,
    exportMode: ExportMode,
): Promise<EffectiveProjectModuleSyncSettings> {
    return updateProjectModuleSyncSettings(projectPath, { exportMode });
}

export function projectSettingsWithModuleSyncPatch(
    existing: ProjectSettingsConfig,
    patch: ProjectModuleSyncSettingsPatch,
): ProjectSettingsConfig {
    return {
        ...existing,
        exportFolder: patch.folderPath ?? existing.exportFolder,
        exportMode: patch.exportMode === undefined
            ? existing.exportMode
            : normalizeExportMode(patch.exportMode),
        importMode: patch.importMode === undefined
            ? existing.importMode
            : normalizeImportMode(patch.importMode),
    };
}
