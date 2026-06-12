import type { ImportMode, ModuleSyncModeSource } from './moduleSyncPlan';
import {
    normalizeExportMode,
    normalizeImportMode,
    readWorkbookSettings,
    resolveWorkbookSetting,
    settingsPathForWorkbook,
    updateWorkbookSettings,
    type ExportMode,
    type WorkbookSettingsConfig,
} from './workbookSettings';

export const DEFAULT_WORKBOOK_EXPORT_MODE: ExportMode = 'exportAll';
export const DEFAULT_WORKBOOK_IMPORT_MODE: ImportMode = 'updateOnly';

export type WorkbookModuleSyncFolderSource = 'workbook' | 'session' | 'missing';
export type WorkbookModuleSyncModeSource = ModuleSyncModeSource;

export interface EffectiveWorkbookModuleSyncSettings {
    folderPath?: string;
    folderPathSource: WorkbookModuleSyncFolderSource;
    exportMode: ExportMode;
    exportModeSource: WorkbookModuleSyncModeSource;
    importMode: ImportMode;
    importModeSource: WorkbookModuleSyncModeSource;
    settingsPath: string;
}

export interface WorkbookModuleSyncSettingsPatch {
    folderPath?: string;
    exportMode?: ExportMode;
    importMode?: ImportMode;
}

export async function effectiveWorkbookModuleSyncSettings(
    workbookPath: string,
): Promise<EffectiveWorkbookModuleSyncSettings> {
    return effectiveWorkbookModuleSyncSettingsFromConfig(
        workbookPath,
        await readWorkbookSettings(workbookPath),
    );
}

export function effectiveWorkbookModuleSyncSettingsFromConfig(
    workbookPath: string,
    config: WorkbookSettingsConfig,
): EffectiveWorkbookModuleSyncSettings {
    const exportMode = resolveWorkbookSetting(config.exportMode, {
        value: DEFAULT_WORKBOOK_EXPORT_MODE,
        source: 'default',
    });
    const importMode = resolveWorkbookSetting(config.importMode, {
        value: DEFAULT_WORKBOOK_IMPORT_MODE,
        source: 'default',
    });
    return {
        folderPath: config.exportFolder,
        folderPathSource: config.exportFolder ? 'workbook' : 'missing',
        exportMode: exportMode.value,
        exportModeSource: exportMode.source,
        importMode: importMode.value,
        importModeSource: importMode.source,
        settingsPath: settingsPathForWorkbook(workbookPath),
    };
}

export async function updateWorkbookModuleSyncSettings(
    workbookPath: string,
    patch: WorkbookModuleSyncSettingsPatch,
): Promise<EffectiveWorkbookModuleSyncSettings> {
    const updated = await updateWorkbookSettings(workbookPath, (existing) =>
        workbookSettingsWithModuleSyncPatch(existing, patch),
    );
    return effectiveWorkbookModuleSyncSettingsFromConfig(workbookPath, updated);
}

export async function setWorkbookModuleSyncExportMode(
    workbookPath: string,
    exportMode: ExportMode,
): Promise<EffectiveWorkbookModuleSyncSettings> {
    return updateWorkbookModuleSyncSettings(workbookPath, { exportMode });
}

export function workbookSettingsWithModuleSyncPatch(
    existing: WorkbookSettingsConfig,
    patch: WorkbookModuleSyncSettingsPatch,
): WorkbookSettingsConfig {
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
