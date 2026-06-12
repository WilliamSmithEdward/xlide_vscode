import * as vscode from 'vscode';
import * as path from 'path';
import type { ImportMode, ModuleSyncFolderSource, ModuleSyncModeSource, ModuleSyncPlan } from './moduleSyncPlan';
import { settingsPathForWorkbook, type ExportMode } from './workbookSettings';
import { measurePerformance, measurePerformanceSync } from './performanceTrace';
import { randomNonce, scriptJson } from './webview/html';
import { webviewHeadHtml } from './webview/page';
import { DebouncedRefresher, RefreshGate } from './webview/refresh';
import { renderWebviewTemplate } from './webview/templates';
import { errorMessage } from './util/errors';

export interface ModuleSyncApplyResult {
    summary: string;
    changed: number;
    skipped: number;
    removed?: number;
    failed: number;
}

export interface ModuleSyncSettings {
    folderPath: string;
    folderPathSource?: ModuleSyncFolderSource;
    exportMode?: ExportMode;
    exportModeSource?: ModuleSyncModeSource;
    importMode?: ImportMode;
    importModeSource?: ModuleSyncModeSource;
    settingsPath?: string;
}

export interface ModuleSyncPreviewOptions {
    onChooseFolder?: (current: ModuleSyncSettings) => Promise<ModuleSyncPlan | undefined>;
    onRefresh?: (settings: ModuleSyncSettings) => Promise<ModuleSyncPlan>;
    onReloadWorkbookSettings?: () => Promise<ModuleSyncPlan | undefined>;
    onSaveSettings?: (settings: ModuleSyncSettings) => Promise<ModuleSyncApplyResult>;
}

export function openModuleSyncPreview(
    context: vscode.ExtensionContext,
    plan: ModuleSyncPlan,
    onApply: (plan: ModuleSyncPlan, selectedIds: readonly string[]) => Promise<ModuleSyncApplyResult>,
    options: ModuleSyncPreviewOptions = {},
): Promise<ModuleSyncApplyResult | undefined> {
    return new Promise((resolve) => {
        const panel = vscode.window.createWebviewPanel(
            'xlide.moduleSyncPreview',
            plan.direction === 'export' ? 'XLIDE Export Preview' : 'XLIDE Import Preview',
            vscode.ViewColumn.One,
            {
                enableScripts: true,
                // Required: plan refreshes arrive as postMessage updates, so the
                // stored html goes stale and the selection state lives in the DOM.
                retainContextWhenHidden: true,
                localResourceRoots: [context.extensionUri],
            },
        );
        let resolved = false;
        let disposed = false;
        let watchedFolderPath: string | undefined;
        let folderWatcherDisposables: vscode.Disposable[] = [];
        const done = (result: ModuleSyncApplyResult | undefined): void => {
            if (!resolved) {
                resolved = true;
                resolve(result);
            }
        };

        let currentPlan = plan;
        const disposeFolderWatcher = (): void => {
            for (const disposable of folderWatcherDisposables) {
                disposable.dispose();
            }
            folderWatcherDisposables = [];
            watchedFolderPath = undefined;
        };
        const postRefreshError = (err: unknown): void => {
            void panel.webview.postMessage({ type: 'error', error: errorMessage(err) });
        };
        const gate = new RefreshGate();
        const folderRefresher = new DebouncedRefresher({
            refresh: () => refreshPlanFromDisk(),
            onError: postRefreshError,
            defaultDelayMs: 300,
            gate,
        });
        const settingsRefresher = new DebouncedRefresher({
            refresh: () => refreshPlanFromWorkbookSettings(),
            onError: postRefreshError,
            defaultDelayMs: 300,
            gate,
        });
        const configureFolderWatcher = (): void => {
            if (!options.onRefresh || disposed || !currentPlan.folderPath || watchedFolderPath === currentPlan.folderPath) {
                return;
            }
            disposeFolderWatcher();
            watchedFolderPath = currentPlan.folderPath;
            const watcher = vscode.workspace.createFileSystemWatcher(
                new vscode.RelativePattern(currentPlan.folderPath, '**/*'),
            );
            const refreshIfRelevant = (uri: vscode.Uri): void => {
                if (isModuleSyncWatchedPath(uri.fsPath)) {
                    folderRefresher.schedule();
                }
            };
            folderWatcherDisposables = [
                watcher.onDidChange(refreshIfRelevant),
                watcher.onDidCreate(refreshIfRelevant),
                watcher.onDidDelete(refreshIfRelevant),
                watcher,
            ];
        };
        const runExclusive = <T>(operation: () => Promise<T>): Promise<T> => gate.runExclusive(operation);
        const updateCurrentPlan = (nextPlan: ModuleSyncPlan): void => {
            currentPlan = nextPlan;
            configureFolderWatcher();
        };
        async function refreshPlanFromDisk(): Promise<void> {
            const refresh = options.onRefresh;
            if (!refresh || disposed) {
                return;
            }
            await measurePerformance('moduleSync.refreshFromDisk', currentPlan.direction, async () => {
                await panel.webview.postMessage({ type: 'refreshing', message: 'Disk changes detected. Refreshing preview...' });
                updateCurrentPlan(await refresh(settingsFromPlan(currentPlan)));
                await panel.webview.postMessage({
                    type: 'plan',
                    plan: currentPlan,
                    message: 'Disk changes detected. Preview refreshed.',
                });
            });
        }
        async function refreshPlanFromWorkbookSettings(): Promise<void> {
            const reload = options.onReloadWorkbookSettings;
            if (!reload || disposed) {
                return;
            }
            await measurePerformance('moduleSync.refreshWorkbookSettings', currentPlan.direction, async () => {
                await panel.webview.postMessage({ type: 'refreshing', message: 'Workbook settings changed. Refreshing preview...' });
                const nextPlan = await reload();
                if (nextPlan) {
                    updateCurrentPlan(nextPlan);
                    await panel.webview.postMessage({
                        type: 'plan',
                        plan: currentPlan,
                        message: 'Workbook settings changed. Preview refreshed.',
                    });
                } else {
                    await panel.webview.postMessage({ type: 'ready' });
                }
            });
        }
        configureFolderWatcher();
        const workbookSettingsPath = settingsPathForWorkbook(currentPlan.workbookPath);
        const workbookSettingsWatcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(
            path.dirname(workbookSettingsPath),
            path.basename(workbookSettingsPath),
        ));
        const workbookSettingsWatcherDisposables = [
            workbookSettingsWatcher.onDidCreate(() => settingsRefresher.schedule()),
            workbookSettingsWatcher.onDidChange(() => settingsRefresher.schedule()),
            workbookSettingsWatcher.onDidDelete(() => settingsRefresher.schedule()),
            workbookSettingsWatcher,
        ];
        panel.webview.html = measurePerformanceSync(
            'moduleSync.renderHtml',
            `${currentPlan.direction}:${currentPlan.items.length} items`,
            () => renderModuleSyncHtml(currentPlan),
        );
        const messageSub = panel.webview.onDidReceiveMessage(async (message: {
            type?: string;
            selectedIds?: string[];
            folderPath?: string;
            exportMode?: ExportMode;
            importMode?: ImportMode;
            itemId?: string;
            side?: 'left' | 'right';
            showHeaders?: boolean;
            autoSaveSettings?: boolean;
            quiet?: boolean;
        }) => {
            if (message.type === 'cancel') {
                done(undefined);
                panel.dispose();
                return;
            }
            if (message.type === 'copy-code') {
                const item = currentPlan.items.find((candidate) => candidate.id === message.itemId);
                if (!item || (message.side !== 'left' && message.side !== 'right')) {
                    await panel.webview.postMessage({ type: 'error', error: 'No code selected to copy.' });
                    return;
                }
                const code = message.side === 'left'
                    ? (message.showHeaders ? item.leftRawCode : item.leftCode)
                    : (message.showHeaders ? item.rightRawCode : item.rightCode);
                await vscode.env.clipboard.writeText(code);
                await panel.webview.postMessage({ type: 'copied', side: message.side });
                return;
            }
            if (message.type === 'choose-folder') {
                const chooseFolder = options.onChooseFolder;
                if (!chooseFolder) {
                    return;
                }
                await panel.webview.postMessage({ type: 'refreshing' });
                try {
                    await measurePerformance('moduleSync.chooseFolder', currentPlan.direction, async () => {
                        await runExclusive(async () => {
                            const nextPlan = await chooseFolder(settingsFromMessage(currentPlan, message));
                            if (nextPlan) {
                                updateCurrentPlan(nextPlan);
                                await panel.webview.postMessage({
                                    type: 'plan',
                                    plan: currentPlan,
                                    message: 'Settings updated. Review the refreshed diff before applying.',
                                    autoSaveSettings: true,
                                });
                            } else {
                                await panel.webview.postMessage({ type: 'ready' });
                            }
                        });
                    });
                } catch (err) {
                    const error = errorMessage(err);
                    await panel.webview.postMessage({ type: 'error', error });
                }
                return;
            }
            if (message.type === 'refresh-settings') {
                const refresh = options.onRefresh;
                if (!refresh) {
                    return;
                }
                await panel.webview.postMessage({ type: 'refreshing' });
                try {
                    await measurePerformance('moduleSync.refreshSettings', currentPlan.direction, async () => {
                        await runExclusive(async () => {
                            updateCurrentPlan(await refresh(settingsFromMessage(currentPlan, message)));
                            await panel.webview.postMessage({
                                type: 'plan',
                                plan: currentPlan,
                                message: 'Settings updated. Review the refreshed diff before applying.',
                                autoSaveSettings: message.autoSaveSettings === true,
                            });
                        });
                    });
                } catch (err) {
                    const error = errorMessage(err);
                    await panel.webview.postMessage({ type: 'error', error });
                }
                return;
            }
            if (message.type === 'save-settings') {
                const saveSettings = options.onSaveSettings;
                if (!saveSettings) {
                    return;
                }
                if (!message.quiet) {
                    await panel.webview.postMessage({ type: 'saving-settings' });
                }
                try {
                    await runExclusive(async () => {
                        const result = await saveSettings(settingsFromMessage(currentPlan, message));
                        await panel.webview.postMessage({ type: 'settings-saved', result, quiet: message.quiet === true });
                    });
                } catch (err) {
                    const error = errorMessage(err);
                    await panel.webview.postMessage({ type: 'error', error });
                }
                return;
            }
            if (message.type !== 'apply') {
                return;
            }
            const selectedIds = message.selectedIds ?? [];
            await panel.webview.postMessage({ type: 'applying' });
            try {
                await measurePerformance('moduleSync.apply', currentPlan.direction, async () => {
                    await runExclusive(async () => {
                        const result = await onApply(currentPlan, selectedIds);
                        await panel.webview.postMessage({ type: 'applied', result });
                        done(result);
                    });
                });
            } catch (err) {
                const error = errorMessage(err);
                await panel.webview.postMessage({ type: 'error', error });
            }
        });
        panel.onDidDispose(() => {
            disposed = true;
            folderRefresher.dispose();
            settingsRefresher.dispose();
            disposeFolderWatcher();
            for (const disposable of workbookSettingsWatcherDisposables) {
                disposable.dispose();
            }
            messageSub.dispose();
            done(undefined);
        });
    });
}

function renderModuleSyncHtml(plan: ModuleSyncPlan): string {
    const nonce = randomNonce();
    return renderWebviewTemplate('assets/webview/moduleSync.html', {
        head: webviewHeadHtml(nonce, plan.title),
        nonce,
        css: renderWebviewTemplate('assets/webview/moduleSync.css', {}),
        js: renderWebviewTemplate('assets/webview/moduleSync.js', {
            planJson: scriptJson(plan),
        }),
    });
}


function settingsFromMessage(
    plan: ModuleSyncPlan,
    message: { folderPath?: string; exportMode?: ExportMode; importMode?: ImportMode },
): ModuleSyncSettings {
    const exportModeSource = message.exportMode !== undefined && message.exportMode !== plan.exportMode
        ? 'session'
        : plan.exportModeSource;
    const importModeSource = message.importMode !== undefined && message.importMode !== plan.importMode
        ? 'session'
        : plan.importModeSource;
    return {
        folderPath: message.folderPath ?? plan.folderPath,
        folderPathSource: plan.folderPathSource,
        exportMode: message.exportMode ?? plan.exportMode,
        exportModeSource,
        importMode: message.importMode ?? plan.importMode,
        importModeSource,
        settingsPath: plan.settingsPath,
    };
}

function settingsFromPlan(plan: ModuleSyncPlan): ModuleSyncSettings {
    return {
        folderPath: plan.folderPath,
        folderPathSource: plan.folderPathSource,
        exportMode: plan.exportMode,
        exportModeSource: plan.exportModeSource,
        importMode: plan.importMode,
        importModeSource: plan.importModeSource,
        settingsPath: plan.settingsPath,
    };
}

function isModuleSyncWatchedPath(filePath: string): boolean {
    return /\.(bas|cls)$/i.test(filePath);
}
