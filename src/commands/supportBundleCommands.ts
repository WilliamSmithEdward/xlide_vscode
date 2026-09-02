import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import {
    activeLocalVbaEditor,
    decodeModuleUri,
} from '../xlideFileSystem';
import {
    analyzeProject,
    projectProblemsForModule,
} from '../vbaProjectWideAnalysis';
import {
    anonymizedProjectAnalysisReportFromResult,
    buildSupportBundle,
    defaultSupportBundleFileName,
    supportBundleDisclosureText,
    supportDiagnosticsText,
    type SupportBundle,
    type SupportBundleAnonymizedAnalysisReport,
    type SupportBundleAnalysisSummary,
    type SupportBundleSetting,
    type SupportBundleProjectSummary,
} from '../supportBundle';
import {
    errorCategoryForSupportLog,
    recentXlideCommands,
} from '../xlideCommandLog';
import { recentXlideOutputLog } from '../xlideOutputLog';
import { recentXlideWriteAudits } from '../xlideWriteAudit';
import { formatPerformanceSnapshot } from '../performanceTrace';
import { resolvedXlideGlobalSettingsFromConfig } from '../globalSettings';
import { registerXlideCommand } from '../xlideCommandRegistration';
import { analyzeOpenModule } from './analysisCommands';
import {
    activeLocalProjectPath,
    statusMessage,
    type CommandDeps,
} from './shared';

interface SupportBundleOptions {
    includeAnonymizedProjectAnalysisReport?: boolean;
    includeSelectedLogs?: boolean;
}

export function registerSupportBundleCommands(deps: CommandDeps): vscode.Disposable[] {
    const { context, bridge, out, vbaIndex } = deps;

    function log(msg: string): void {
        out.appendLine(msg);
    }

    async function activeModuleSupportData(): Promise<{
        project: SupportBundleProjectSummary;
        analysis: SupportBundleAnalysisSummary;
    }> {
        const editor = activeLocalVbaEditor();
        if (!editor) {
            return {
                project: { available: false },
                analysis: { available: false },
            };
        }

        const { projectPath, moduleName } = decodeModuleUri(editor.document.uri);
        const source = editor.document.getText();
        const { modules, moduleType, result } = await analyzeOpenModule(vbaIndex, projectPath, moduleName, source);
        const moduleTypes = countBy(modules.map((mod) => mod.type || 'unknown'));
        const project: SupportBundleProjectSummary = {
            available: true,
            projectPath: projectPath,
            extension: path.extname(projectPath).toLowerCase(),
            moduleCount: modules.length,
            moduleTypes,
            activeModuleType: moduleType,
        };

        const problems = projectProblemsForModule(
            moduleName,
            moduleType,
            source,
            result.diagnostics,
        );
        return {
            project,
            analysis: {
                available: true,
                moduleType,
                errorCount: problems.filter((problem) => problem.severity === 'error').length,
                warningCount: problems.filter((problem) => problem.severity === 'warning').length,
                suppressedCount: result.suppressedCount,
                byCode: countBy(problems.map((problem) => problem.code || 'unclassified')),
            },
        };
    }

    function countBy(values: readonly string[]): Record<string, number> {
        const out: Record<string, number> = {};
        for (const value of values) {
            out[value] = (out[value] ?? 0) + 1;
        }
        return out;
    }

    function xlideSettingsForSupportBundle(): SupportBundleSetting[] {
        return resolvedXlideGlobalSettingsFromConfig(vscode.workspace.getConfiguration('xlide'));
    }

    async function anonymizedProjectAnalysisReportForActiveProject():
        Promise<SupportBundleAnonymizedAnalysisReport> {
        const projectPath = await activeLocalProjectPath();
        if (!projectPath) {
            return { included: false, unavailableReason: 'no-active-project' };
        }
        try {
            return anonymizedProjectAnalysisReportFromResult(await analyzeProject(bridge, projectPath));
        } catch (err) {
            return {
                included: false,
                unavailableReason: 'analysis-failed',
                errorCategory: errorCategoryForSupportLog(err),
            };
        }
    }

    async function currentSupportBundle(
        now = new Date(),
        options: SupportBundleOptions = {},
    ): Promise<SupportBundle> {
        const packageJson = context.extension.packageJSON as {
            name?: string;
            publisher?: string;
            version?: string;
            displayName?: string;
        };
        const active = await activeModuleSupportData();
        const anonymizedProjectAnalysisReport = options.includeAnonymizedProjectAnalysisReport
            ? await anonymizedProjectAnalysisReportForActiveProject()
            : undefined;
        return buildSupportBundle({
            generatedAt: now.toISOString(),
            extension: {
                id: packageJson.publisher && packageJson.name
                    ? `${packageJson.publisher}.${packageJson.name}`
                    : packageJson.name,
                name: packageJson.displayName ?? packageJson.name,
                version: packageJson.version,
            },
            vscode: {
                version: vscode.version,
                appName: vscode.env.appName,
            },
            runtime: {
                platform: process.platform,
                arch: process.arch,
                node: process.version,
            },
            workspace: {
                folderCount: vscode.workspace.workspaceFolders?.length ?? 0,
            },
            settings: xlideSettingsForSupportBundle(),
            project: active.project,
            analysis: active.analysis,
            commands: recentXlideCommands(),
            writeAudits: recentXlideWriteAudits(),
            anonymizedProjectAnalysisReport,
            selectedLogs: options.includeSelectedLogs ? recentXlideOutputLog() : undefined,
        });
    }

    async function selectSupportBundleExportOptions(
        bundle: SupportBundle,
    ): Promise<SupportBundleOptions | undefined> {
        const choice = await vscode.window.showInformationMessage(
            'XLIDE support bundle export',
            {
                modal: true,
                detail: supportBundleDisclosureText(bundle),
            },
            'Export',
            'Choose Extras',
            'Copy Diagnostics',
        );
        if (choice === 'Copy Diagnostics') {
            await copyDiagnosticsFromBundle(bundle);
            return undefined;
        }
        if (choice === 'Choose Extras') {
            const picks = await vscode.window.showQuickPick(
                [
                    {
                        label: 'Anonymized file analysis report',
                        description: 'Counts by rule/module type only; no source or module names',
                        picked: true,
                        option: 'includeAnonymizedProjectAnalysisReport' as const,
                    },
                    {
                        label: 'Selected recent XLIDE logs',
                        description: 'Recent XLIDE-authored output lines with paths redacted',
                        picked: false,
                        option: 'includeSelectedLogs' as const,
                    },
                ],
                {
                    title: 'XLIDE Support Bundle: Optional Extras',
                    canPickMany: true,
                    placeHolder: 'Choose only the extras you want included in this export.',
                },
            );
            if (!picks) {
                return undefined;
            }
            return {
                includeAnonymizedProjectAnalysisReport:
                    picks.some((pick) => pick.option === 'includeAnonymizedProjectAnalysisReport'),
                includeSelectedLogs:
                    picks.some((pick) => pick.option === 'includeSelectedLogs'),
            };
        }
        return choice === 'Export' ? {} : undefined;
    }

    async function copyDiagnosticsFromBundle(bundle: SupportBundle): Promise<void> {
        await vscode.env.clipboard.writeText(supportDiagnosticsText(bundle));
        statusMessage('XLIDE: Redacted diagnostics copied to clipboard.');
    }

    async function copyDiagnostics(): Promise<void> {
        await copyDiagnosticsFromBundle(await currentSupportBundle());
    }

    async function exportSupportBundle(): Promise<void> {
        const now = new Date();
        const baseBundle = await currentSupportBundle(now);
        const options = await selectSupportBundleExportOptions(baseBundle);
        if (!options) {
            return;
        }
        const bundle = Object.keys(options).length === 0
            ? baseBundle
            : await currentSupportBundle(now, options);

        const defaultFolder = vscode.workspace.workspaceFolders?.[0]?.uri;
        const target = await vscode.window.showSaveDialog({
            title: 'Export XLIDE Support Bundle (redacted JSON; no VBA source)',
            defaultUri: defaultFolder
                ? vscode.Uri.joinPath(defaultFolder, defaultSupportBundleFileName(now))
                : undefined,
            filters: { JSON: ['json'] },
        });
        if (!target) {
            return;
        }

        await fs.promises.writeFile(target.fsPath, `${JSON.stringify(bundle, null, 2)}\n`, 'utf8');
        // The user chose the destination in the save dialog moments ago.
        statusMessage(`XLIDE: Support bundle exported to ${target.fsPath}`);
    }

    return [
        // Export a redacted local diagnostic snapshot for support/self-debugging.
        registerXlideCommand('xlide.exportSupportBundle', () => exportSupportBundle(), {
            errorPrefix: 'Failed to export support bundle',
            logTag: 'supportBundle',
            log,
        }),

        registerXlideCommand('xlide.copyDiagnostics', () => copyDiagnostics(), {
            errorPrefix: 'Failed to copy diagnostics',
            logTag: 'copyDiagnostics',
            log,
        }),

        registerXlideCommand('xlide.copyPerformanceSnapshot', async () => {
            await vscode.env.clipboard.writeText(formatPerformanceSnapshot());
            vscode.window.showInformationMessage('XLIDE: Performance snapshot copied to clipboard.');
        }, {
            errorPrefix: 'Failed to copy performance snapshot',
            logTag: 'copyPerformanceSnapshot',
            log,
        }),
    ];
}
