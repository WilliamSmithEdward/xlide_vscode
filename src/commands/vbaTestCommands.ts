import * as vscode from 'vscode';
import * as path from 'path';
import {
    encodeModuleUri,
    decodeModuleUri,
    sameWorkbookPath,
    XLIDE_VBA_LANGUAGE_ID,
    activeLocalVbaEditor,
    workbookIdentityKey,
} from '../xlideFileSystem';
import {
    describeVbaTestSelection,
    discoverWorkbookVbaTests,
    summarizeVbaTestTags,
    summarizeVbaTestRun,
    type VbaTestCase,
    type VbaTestRunItem,
    type VbaTestRunReport,
    type VbaTestSelectionOptions,
} from '../vbaTestRunner';
import { type VbaTestRunOptions } from '../vbaTestExecution';
import { executeVbaTestRun } from '../vbaTestRunPipeline';
import { openVbaTestResults, setVbaTestResultsRunning } from '../vbaTestResultsWebview';
import { checkExcelComAvailability } from '../excelComAvailability';
import {
    openVbaTestsPanel,
    type VbaTestsRunFilterRequest,
    type VbaTestsRunSelectedRequest,
    type VbaTestSupportStatusModel,
    type VbaTestsPanelModel,
} from '../vbaTestsWebview';
import {
    normalizeVbaTestSupportModuleSource,
    XLIDE_ASSERT_MODULE_NAME,
    XLIDE_ASSERT_MODULE_SOURCE,
} from '../vbaTestSupportModule';
import { getVbaTestSupportStatus } from '../vbaTestSupportStatus';
import { registerXlideCommand } from '../xlideCommandRegistration';
import { recordXlideWriteAuditEvent as recordWriteAudit } from '../xlideWriteAudit';
import { writeWorkbookModule } from '../workbookModuleOperations';
import type { XlideNode } from '../xlsmExplorer';
import { errorMessage } from '../util/errors';
import {
    logChangeSummary,
    procedureNameAtCursor,
    resolveWorkbookPath,
    showAnalysisSourceDocument,
    type CommandDeps,
} from './shared';

interface VbaTestLastFailedRun {
    testIds: string[];
    tests: Array<{
        id: string;
        qualifiedName: string;
        status: VbaTestRunItem['status'];
    }>;
}

const VBA_TEST_RERUN_FAILED_STATUSES = new Set<VbaTestRunItem['status']>([
    'failed',
    'timeout',
    'host-error',
    'xpass',
]);

export function registerVbaTestCommands(deps: CommandDeps): vscode.Disposable[] {
    const { context, bridge, explorer, out } = deps;
    const lastFailedVbaTestRuns = new Map<string, VbaTestLastFailedRun>();

    function log(msg: string): void {
        out.appendLine(msg);
    }

    function updateLastFailedVbaTestRun(report: VbaTestRunReport): void {
        const failed = report.results
            .filter((result) => VBA_TEST_RERUN_FAILED_STATUSES.has(result.status))
            .map((result) => ({
                id: result.test.id,
                qualifiedName: result.test.qualifiedName,
                status: result.status,
            }));
        const key = workbookIdentityKey(report.filePath);
        if (failed.length === 0) {
            lastFailedVbaTestRuns.delete(key);
            return;
        }
        lastFailedVbaTestRuns.set(key, {
            testIds: failed.map((test) => test.id),
            tests: failed,
        });
    }

    function lastFailedRunForWorkbook(filePath: string): VbaTestLastFailedRun | undefined {
        return lastFailedVbaTestRuns.get(workbookIdentityKey(filePath));
    }

    async function rerunFailedVbaTestsForWorkbook(filePath: string): Promise<void> {
        const failed = lastFailedRunForWorkbook(filePath);
        if (!failed || failed.testIds.length === 0) {
            void vscode.window.showInformationMessage(
                `XLIDE: No failed VBA tests to rerun for "${path.basename(filePath)}".`,
            );
            return;
        }
        await runVbaTestsForWorkbook(filePath, {
            selection: {
                testIds: failed.testIds,
            },
        });
    }

    async function openVbaTestCase(filePath: string, test: VbaTestCase): Promise<void> {
        const uri = encodeModuleUri(filePath, test.moduleName);
        const doc = await vscode.workspace.openTextDocument(uri);
        await vscode.languages.setTextDocumentLanguage(doc, XLIDE_VBA_LANGUAGE_ID);
        const editor = await showAnalysisSourceDocument(doc);
        const line = Math.max(0, test.line - 1);
        const column = Math.max(0, test.column - 1);
        const position = new vscode.Position(line, column);
        editor.selection = new vscode.Selection(position, position);
        editor.revealRange(new vscode.Range(position, position), vscode.TextEditorRevealType.InCenterIfOutsideViewport);
    }

    async function runVbaTestsForWorkbook(
        filePath: string,
        options: VbaTestRunOptions = {},
    ): Promise<void> {
        const name = path.basename(filePath);
        const selectionDescription = describeVbaTestSelection(options.selection);
        const runScope = selectionDescription ? ` (${selectionDescription})` : '';
        let resultsPanelRunning = false;
        try {
            const result = await executeVbaTestRun(bridge, filePath, {
                ...options,
                log,
                runTests: (run) => {
                    resultsPanelRunning = true;
                    setVbaTestResultsRunning(filePath, true);
                    return vscode.window.withProgress(
                        {
                            location: vscode.ProgressLocation.Notification,
                            title: `XLIDE: Running VBA tests${runScope} for "${name}"...`,
                            cancellable: false,
                        },
                        (progress) => run(progress),
                    );
                },
            });
            if (result.kind === 'blocked-support') {
                openVbaTestsForWorkbook(filePath);
                void vscode.window.showWarningMessage(
                    `XLIDE: Install XlideAssert.bas from the Unit Tests GUI before running tests for "${name}".`,
                );
                return;
            }
            if (result.kind === 'blocked-com') {
                openVbaTestsForWorkbook(filePath);
                void vscode.window.showWarningMessage(`XLIDE: ${result.runtime.description}`);
                return;
            }
            const { report } = result.execution;
            log(`[runVbaTests] Report JSON:\n${JSON.stringify(report, null, 2)}`);
            if (result.artifacts.ok) {
                const { artifacts, settings } = result.artifacts;
                log(`[runVbaTests] Artifacts written to ${artifacts.runDirectory}`);
                log(`[runVbaTests] CI status written to ${artifacts.statusPath}`);
                log(`[runVbaTests] Artifact settings source folder=${settings.artifactFolderSource} retention=${settings.artifactRetentionSource}`);
            } else {
                log(`[runVbaTests] Artifact write failed: ${result.artifacts.error}`);
                void vscode.window.showWarningMessage(`XLIDE: VBA test artifacts could not be written: ${result.artifacts.error}`);
            }
            updateLastFailedVbaTestRun(report);
            openVbaTestResults(context, report, {
                onRerunFailed: () => rerunFailedVbaTestsForWorkbook(filePath),
                onOpenTest: (test) => openVbaTestCase(filePath, test),
            });
            showVbaTestRunOutcome(report);
        } catch (err) {
            const msg = errorMessage(err);
            log(`[runVbaTests] FAILED: ${msg}`);
            vscode.window.showErrorMessage(`XLIDE: VBA tests failed: ${msg}`);
        } finally {
            if (resultsPanelRunning) {
                setVbaTestResultsRunning(filePath, false);
            }
        }
    }

    function openVbaTestsForWorkbook(filePath: string): void {
        openVbaTestsPanel(context, filePath, {
            getModel: () => vbaTestsPanelModel(filePath),
            onInstallSupport: async () => {
                await installVbaTestSupportModule(filePath);
            },
            onRunAll: async () => {
                await runVbaTestsForWorkbook(filePath);
            },
            onRunWithFilters: async (filters) => {
                await runVbaTestsForWorkbook(filePath, vbaTestRunOptionsFromFilters(filters));
            },
            onRunSelected: async (selection) => {
                await runSelectedVbaTestsForWorkbook(filePath, selection);
            },
            onRunCurrentModule: async (request) => {
                await runCurrentModuleVbaTestsForWorkbook(filePath, request.failFast);
            },
            onRunCurrentTest: async (request) => {
                await runCurrentTestForWorkbook(filePath, request.failFast);
            },
            onRerunFailed: async () => {
                await rerunFailedVbaTestsForWorkbook(filePath);
            },
            onDidChangeWorkbookTree: explorer.onDidChangeTreeData,
        });
    }

    async function vbaTestsPanelModel(filePath: string): Promise<VbaTestsPanelModel> {
        const [support, runtime, discovery] = await Promise.all([
            vbaTestSupportStatus(filePath),
            checkExcelComAvailability(),
            vbaTestsDiscoveryStatus(filePath),
        ]);
        const lastFailed = lastFailedRunForWorkbook(filePath);
        return {
            filePath,
            workbookName: path.basename(filePath),
            support,
            runtime,
            discovery,
            lastFailed: lastFailed
                ? {
                    count: lastFailed.testIds.length,
                    tests: lastFailed.tests,
                }
                : undefined,
        };
    }

    async function vbaTestsDiscoveryStatus(filePath: string): Promise<VbaTestsPanelModel['discovery']> {
        try {
            const discovery = await discoverWorkbookVbaTests(bridge, filePath);
            const tags = summarizeVbaTestTags(discovery.tests);
            const taggedTests = discovery.tests.filter((test) => test.metadata.tags.length > 0).length;
            return {
                totalTests: discovery.tests.length,
                taggedTests,
                untaggedTests: Math.max(0, discovery.tests.length - taggedTests),
                tags,
                tests: discovery.tests.map((test) => ({
                    id: test.id,
                    qualifiedName: test.qualifiedName,
                    moduleName: test.moduleName,
                    procedureName: test.procedureName,
                    line: test.line,
                    tags: test.metadata.tags,
                })),
            };
        } catch (err) {
            const error = errorMessage(err);
            return {
                totalTests: 0,
                taggedTests: 0,
                untaggedTests: 0,
                tags: [],
                tests: [],
                error,
            };
        }
    }

    function vbaTestRunOptionsFromFilters(filters: VbaTestsRunFilterRequest): VbaTestRunOptions {
        const selection: VbaTestSelectionOptions = {};
        if (filters.includeTags.length > 0) {
            selection.includeTags = filters.includeTags;
        }
        if (filters.excludeTags.length > 0) {
            selection.excludeTags = filters.excludeTags;
        }
        return {
            selection,
            failFast: filters.failFast,
        };
    }

    async function runSelectedVbaTestsForWorkbook(
        filePath: string,
        request: VbaTestsRunSelectedRequest,
    ): Promise<void> {
        if (request.testIds.length === 0) {
            void vscode.window.showInformationMessage(
                `XLIDE: Select at least one VBA test to run for "${path.basename(filePath)}".`,
            );
            return;
        }
        await runVbaTestsForWorkbook(filePath, {
            selection: {
                testIds: request.testIds,
            },
            failFast: request.failFast,
        });
    }

    async function vbaTestSupportStatus(filePath: string): Promise<VbaTestSupportStatusModel> {
        return getVbaTestSupportStatus(bridge, filePath);
    }

    function showVbaTestRunOutcome(report: VbaTestRunReport): void {
        const summary = summarizeVbaTestRun(report);
        const label = path.basename(report.filePath);
        if (summary.total === 0) {
            const selectionDescription = describeVbaTestSelection(report.discovery.selection);
            if (selectionDescription && report.discovery.unfilteredTestCount > 0) {
                void vscode.window.showInformationMessage(
                    `XLIDE: No VBA tests matched ${selectionDescription} in "${label}".`,
                );
                return;
            }
            void vscode.window.showInformationMessage(
                `XLIDE: No VBA tests were discovered in "${label}". Add '@xlide-test' above a no-argument standard-module Sub.`,
            );
            return;
        }
        if (summary.failed > 0 || summary.xpass > 0 || summary.timeout > 0 || summary.hostError > 0) {
            void vscode.window.showWarningMessage(
                `XLIDE: "${label}" VBA tests finished with ${summary.failed} failed, ${summary.timeout} timed out, ${summary.hostError} host error(s), ${summary.xpass} unexpected passed, ${summary.passed} passed, ${summary.skipped} skipped.`,
            );
            return;
        }
        if (summary.skipped > 0) {
            void vscode.window.showWarningMessage(
                `XLIDE: "${label}" VBA tests were discovered but ${summary.skipped} were skipped.`,
            );
            return;
        }
        void vscode.window.showInformationMessage(
            `XLIDE: "${label}" passed ${summary.passed} VBA test(s).`,
        );
    }

    async function installVbaTestSupportModule(filePath: string): Promise<boolean> {
        const modules = await bridge.call<Array<{ name: string; type: string }>>(
            'listModules',
            { path: filePath },
        );
        const existing = modules.find(
            (module) => module.name.toLowerCase() === XLIDE_ASSERT_MODULE_NAME.toLowerCase(),
        );
        if (existing && existing.type !== 'standard') {
            void vscode.window.showErrorMessage(
                `XLIDE: "${XLIDE_ASSERT_MODULE_NAME}" already exists as a ${existing.type} module. Rename it before installing the test support module.`,
            );
            return false;
        }
        if (existing) {
            const current = await bridge.call<{ source: string }>(
                'readModule',
                { path: filePath, module: existing.name },
            );
            if (
                normalizeVbaTestSupportModuleSource(current.source) ===
                normalizeVbaTestSupportModuleSource(XLIDE_ASSERT_MODULE_SOURCE)
            ) {
                void vscode.window.showInformationMessage(
                    `XLIDE: "${XLIDE_ASSERT_MODULE_NAME}" is already installed in "${path.basename(filePath)}".`,
                );
                return false;
            }
            const choice = await vscode.window.showWarningMessage(
                `Update the existing "${XLIDE_ASSERT_MODULE_NAME}" module in "${path.basename(filePath)}"?`,
                { modal: true },
                'Update',
            );
            if (choice !== 'Update') {
                return false;
            }
        }

        await writeWorkbookModule(deps, {
            filePath,
            moduleName: XLIDE_ASSERT_MODULE_NAME,
            source: XLIDE_ASSERT_MODULE_SOURCE,
            kind: 'standard',
        });
        const summaryText = logChangeSummary(log, 'installVbaTestSupport', {
            operation: existing ? 'Update VBA test support module' : 'Install VBA test support module',
            changed: [XLIDE_ASSERT_MODULE_NAME],
        });
        recordWriteAudit({
            command: 'xlide.installVbaTestSupport',
            operation: 'write-module',
            outcome: 'succeeded',
            workbookPath: filePath,
            moduleName: XLIDE_ASSERT_MODULE_NAME,
            summary: summaryText,
        });
        void vscode.window.showInformationMessage(
            `XLIDE: "${XLIDE_ASSERT_MODULE_NAME}" ${existing ? 'updated' : 'installed'} in "${path.basename(filePath)}".`,
        );
        return true;
    }

    async function runCurrentModuleVbaTestsForWorkbook(filePath: string, failFast = false): Promise<void> {
        const active = await activeVbaTestEditorContext(filePath);
        if (!active) {
            return;
        }
        await runVbaTestsForWorkbook(filePath, {
            selection: { moduleName: active.moduleName },
            failFast,
        });
    }

    async function runCurrentTestForWorkbook(filePath: string, failFast = false): Promise<void> {
        const active = await activeVbaTestEditorContext(filePath);
        if (!active) {
            return;
        }
        const procedureName = procedureNameAtCursor(active.editor);
        if (!procedureName) {
            vscode.window.showWarningMessage('XLIDE: Cursor is not inside a VBA procedure.');
            return;
        }
        await runVbaTestsForWorkbook(filePath, {
            selection: {
                moduleName: active.moduleName,
                procedureName,
            },
            failFast,
        });
    }

    async function activeVbaTestEditorContext(expectedWorkbookPath?: string): Promise<{
        editor: vscode.TextEditor;
        xlsmPath: string;
        moduleName: string;
    } | undefined> {
        const editor = activeLocalVbaEditor();
        if (!editor) {
            vscode.window.showWarningMessage('XLIDE: Open a local workbook VBA module to run module tests.');
            return undefined;
        }
        let xlsmPath: string;
        let moduleName: string;
        try {
            ({ xlsmPath, moduleName } = decodeModuleUri(editor.document.uri));
        } catch {
            // isLocalXlideDocument verifies the scheme but not the *.bas module
            // shape decodeModuleUri requires, so guard against a non-module URI.
            vscode.window.showWarningMessage('XLIDE: Open a local workbook VBA module to run module tests.');
            return undefined;
        }
        if (expectedWorkbookPath && !sameWorkbookPath(xlsmPath, expectedWorkbookPath)) {
            vscode.window.showWarningMessage(
                `XLIDE: Open a VBA module from "${path.basename(expectedWorkbookPath)}" before running current-scope tests from this panel.`,
            );
            return undefined;
        }
        if (editor.document.isDirty) {
            const saved = await editor.document.save();
            if (!saved) {
                vscode.window.showWarningMessage('XLIDE: Save the current module before running VBA tests.');
                return undefined;
            }
        }
        return { editor, xlsmPath, moduleName };
    }

    return [
        // Open the workbook-scoped VBA tests GUI.
        registerXlideCommand('xlide.runVbaTests', async (node: XlideNode) => {
            const filePath = resolveWorkbookPath(node);
            if (!filePath) {
                vscode.window.showWarningMessage('XLIDE: No workbook selected to test.');
                return;
            }
            openVbaTestsForWorkbook(filePath);
        }),
    ];
}
