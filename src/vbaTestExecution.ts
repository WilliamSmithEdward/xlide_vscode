import * as cp from 'child_process';
import * as fs from 'fs';
import * as path from 'path';
import type { PythonBridge } from './pythonBridge';
import {
    buildOwnedReadOnlyExcelTestHostScript,
    buildVbaTestDirectRunnerModule,
    DEFAULT_VBA_TEST_TIMEOUT_MS,
    parseVbaTestHostEventLine,
    vbaTestHostPlanItems,
} from './vbaTestExcelHost';
import {
    validateVbaTestHostOracleTrace,
    type VbaTestHostOracleEvent,
} from './vbaTestHostOracle';
import { createVbaTestHostTempDir } from './vbaTestTempFiles';
import {
    XLIDE_ASSERT_MODULE_NAME,
    XLIDE_ASSERT_MODULE_SOURCE,
} from './vbaTestSupportModule';
import {
    createVbaTestRunReport,
    describeVbaTestSelection,
    discoverWorkbookVbaTests,
    type VbaTestCase,
    type VbaTestRunItem,
    type VbaTestRunReport,
    type VbaTestSelectionOptions,
} from './vbaTestRunner';
import { vbaTestFailureMessage } from './vbaTestFailureMessages';
import { measurePerformance } from './performanceTrace';
import { errorMessage } from './util/errors';
import { runPowerShell } from './util/powershell';

export interface VbaTestRunOptions {
    selection?: VbaTestSelectionOptions;
    failFast?: boolean;
}

export interface VbaTestProgressReporter {
    report(value: { message?: string; increment?: number }): void;
}

export interface RunWorkbookVbaTestsOptions extends VbaTestRunOptions {
    progress?: VbaTestProgressReporter;
    log?: (message: string) => void;
}

export interface VbaTestRunExecution {
    report: VbaTestRunReport;
    hostEvents: VbaTestHostOracleEvent[];
}

interface OwnedReadOnlyExcelHostRunResult {
    events: VbaTestHostOracleEvent[];
    resultsByName: Map<string, OwnedReadOnlyExcelHostTestResult>;
    hostError?: string;
    timedOutAfter?: string;
}

export interface OwnedReadOnlyExcelHostTestResult {
    outcome: 'passed' | 'failed' | 'timeout' | 'modal-blocked' | 'runner-error';
    durationMs: number;
    message?: string;
    errorNumber?: number;
    errorSource?: string;
    output?: string[];
}

type OwnedExcelKillReason = 'timeout' | 'hung' | 'modal-blocked' | 'runner-error' | 'cleanup-failed';

const DEFAULT_VBA_TEST_CLEANUP_GRACE_MS = 5000;

export async function runWorkbookVbaTests(
    bridge: PythonBridge,
    filePath: string,
    options: RunWorkbookVbaTestsOptions = {},
): Promise<VbaTestRunExecution> {
    return measurePerformance('vbaTests.runWorkbook', path.basename(filePath), async () => {
    const log = options.log ?? (() => { /* optional caller logging */ });
    const startedAt = new Date();
    const startedMs = Date.now();
    const runOptions: VbaTestRunOptions = {
        selection: options.selection,
        failFast: options.failFast,
    };
    options.progress?.report({ message: 'Discovering tests...' });
    const discovery = await discoverWorkbookVbaTests(bridge, filePath, runOptions.selection);
    const results: VbaTestRunItem[] = [];

    if (discovery.tests.length === 0) {
        return {
            report: createVbaTestRunReport({
                filePath,
                startedAt,
                durationMs: Date.now() - startedMs,
                discovery,
                results,
            }),
            hostEvents: [],
        };
    }

    if (process.platform !== 'win32') {
        const message = 'VBA test execution currently requires Excel COM on Windows.';
        for (const test of discovery.tests) {
            results.push({
                test,
                status: 'skipped',
                durationMs: 0,
                error: message,
            });
        }
        return {
            report: createVbaTestRunReport({
                filePath,
                startedAt,
                durationMs: Date.now() - startedMs,
                discovery,
                results,
            }),
            hostEvents: [],
        };
    }

    log('[runVbaTests] attachToRunningExcel=false (owned read-only test host)');
    log(`[runVbaTests] selection=${describeVbaTestSelection(runOptions.selection) || 'all tests'} failFast=${Boolean(runOptions.failFast)}`);
    const executableTests = discovery.tests.filter((test) => !test.metadata.skipReason);
    options.progress?.report({ message: `Running ${executableTests.length} test(s) in owned Excel...` });
    const hostRun: OwnedReadOnlyExcelHostRunResult = executableTests.length > 0
        ? await runOwnedReadOnlyExcelTestHost(bridge, filePath, executableTests, runOptions, log)
        : {
            events: [],
            resultsByName: new Map<string, OwnedReadOnlyExcelHostTestResult>(),
        };
    let stoppedAfter: string | undefined;

    for (const test of discovery.tests) {
        if (stoppedAfter) {
            const message = `Not run because fail-fast stopped after ${stoppedAfter}.`;
            results.push({ test, status: 'skipped', durationMs: 0, error: message });
            log(`[runVbaTests] SKIP ${test.qualifiedName}: ${message}`);
            continue;
        }

        if (test.metadata.skipReason) {
            results.push({
                test,
                status: 'skipped',
                durationMs: 0,
                error: test.metadata.skipReason,
            });
            log(`[runVbaTests] SKIP ${test.qualifiedName}: ${test.metadata.skipReason}`);
            continue;
        }

        const hostResult = hostRun.resultsByName.get(test.qualifiedName);
        let result: VbaTestRunItem;
        if (hostResult) {
            result = vbaTestRunItemFromHostResult(test, hostResult);
        } else if (hostRun.hostError) {
            result = { test, status: 'host-error', durationMs: 0, error: hostRun.hostError };
        } else if (hostRun.timedOutAfter) {
            result = {
                test,
                status: 'skipped',
                durationMs: 0,
                error: `Not run because test host timed out after ${hostRun.timedOutAfter}.`,
            };
        } else {
            result = {
                test,
                status: 'host-error',
                durationMs: 0,
                error: 'The Excel test host did not emit a result for this test.',
            };
        }
        results.push(result);
        logVbaTestRunItem(result, log);
        if (shouldFailFastStop(result, runOptions.failFast)) {
            stoppedAfter = test.qualifiedName;
        }
    }

    return {
        report: createVbaTestRunReport({
            filePath,
            startedAt,
            durationMs: Date.now() - startedMs,
            discovery,
            results,
        }),
        hostEvents: hostRun.events,
    };
    });
}

async function runOwnedReadOnlyExcelTestHost(
    bridge: PythonBridge,
    filePath: string,
    tests: readonly VbaTestCase[],
    options: VbaTestRunOptions,
    log: (message: string) => void,
): Promise<OwnedReadOnlyExcelHostRunResult> {
    return measurePerformance('vbaTests.ownedExcelHost', `${path.basename(filePath)} ${tests.length} tests`, async () => {
    const hostScriptDir = await createVbaTestHostTempDir();
    const tempWorkbookPath = path.join(hostScriptDir, path.basename(filePath));
    const hostScriptPath = path.join(hostScriptDir, 'run-vba-tests.ps1');
    const runnerModuleName = `XlideRun${Date.now().toString(36).slice(-8)}`;
    try {
        await fs.promises.copyFile(filePath, tempWorkbookPath);
        await bridge.call<{ ok?: boolean; signatureDropped?: boolean }>('writeModule', {
            path: tempWorkbookPath,
            module: XLIDE_ASSERT_MODULE_NAME,
            source: XLIDE_ASSERT_MODULE_SOURCE,
            kind: 'standard',
        });
        await bridge.call<{ ok?: boolean; signatureDropped?: boolean }>('writeModule', {
            path: tempWorkbookPath,
            module: runnerModuleName,
            source: buildVbaTestDirectRunnerModule(tests, runnerModuleName),
            kind: 'standard',
        });
        const script = buildOwnedReadOnlyExcelTestHostScript(
            tempWorkbookPath,
            vbaTestHostPlanItems(tests),
            { failFast: options.failFast, runnerModuleName },
        );
        await fs.promises.writeFile(hostScriptPath, script, 'utf8');
    } catch (err) {
        try {
            await fs.promises.rm(hostScriptDir, { recursive: true, force: true });
        } catch {
            // Best-effort cleanup only.
        }
        throw err;
    }
    log(`[runVbaTests] Running owned read-only Excel host for ${tests.length} test(s).`);
    log(`[runVbaTests] Temporary workbook path: ${tempWorkbookPath}`);
    log(`[runVbaTests] Host script path: ${hostScriptPath}`);

    return new Promise<OwnedReadOnlyExcelHostRunResult>((resolve) => {
        const events: VbaTestHostOracleEvent[] = [];
        let currentMacro: { excelId: string; qualifiedName: string; timeoutMs: number; startedMs: number } | undefined;
        let currentModalBlocker: Extract<VbaTestHostOracleEvent, { kind: 'modal-blocked' }> | undefined;
        let currentTimer: ReturnType<typeof setTimeout> | undefined;
        let startupTimer: ReturnType<typeof setTimeout> | undefined;
        let cleanupTimer: ReturnType<typeof setTimeout> | undefined;
        let ownedExcelPid: number | undefined;
        let ownedExcelKilled = false;
        let sawWorkbookClosed = false;
        let sawExcelQuit = false;
        let timedOutAfter: string | undefined;
        let settled = false;
        let hostScriptCleaned = false;

        const clearCurrentTimer = () => {
            if (currentTimer) {
                clearTimeout(currentTimer);
                currentTimer = undefined;
            }
        };
        const clearStartupTimer = () => {
            if (startupTimer) {
                clearTimeout(startupTimer);
                startupTimer = undefined;
            }
        };
        const clearCleanupTimer = () => {
            if (cleanupTimer) {
                clearTimeout(cleanupTimer);
                cleanupTimer = undefined;
            }
        };
        const cleanupHostScript = () => {
            if (hostScriptCleaned) {
                return;
            }
            hostScriptCleaned = true;
            void fs.promises.rm(hostScriptDir, { recursive: true, force: true }).catch((err) => {
                const message = errorMessage(err);
                log(`[runVbaTests] Could not delete temporary host script: ${message}`);
                setTimeout(() => {
                    void fs.promises.rm(hostScriptDir, { recursive: true, force: true }).catch(() => {
                        // Best-effort cleanup only.
                    });
                }, 1000);
            });
        };

        const eventResult = (event: VbaTestHostOracleEvent): OwnedReadOnlyExcelHostTestResult | undefined => {
            if (event.kind !== 'macro-finished') {
                return undefined;
            }
            if (event.outcome === 'passed' || event.outcome === 'failed') {
                return {
                    outcome: event.outcome,
                    durationMs: event.durationMs ?? 0,
                    message: event.message,
                    errorNumber: event.errorNumber,
                    errorSource: event.errorSource,
                    output: event.output,
                };
            }
            if (event.outcome === 'timeout' || event.outcome === 'hung') {
                return {
                    outcome: 'timeout',
                    durationMs: event.durationMs ?? currentMacro?.timeoutMs ?? DEFAULT_VBA_TEST_TIMEOUT_MS,
                    message: event.message,
                };
            }
            if (event.outcome === 'modal-blocked') {
                return {
                    outcome: 'modal-blocked',
                    durationMs: event.durationMs ?? currentMacro?.timeoutMs ?? DEFAULT_VBA_TEST_TIMEOUT_MS,
                    message: event.message,
                };
            }
            return {
                outcome: 'runner-error',
                durationMs: event.durationMs ?? 0,
                message: event.message,
            };
        };

        const finish = (hostError?: string) => {
            if (settled) {
                return;
            }
            settled = true;
            clearCurrentTimer();
            clearStartupTimer();
            clearCleanupTimer();
            cleanupHostScript();
            const resultsByName = new Map<string, OwnedReadOnlyExcelHostTestResult>();
            for (const event of events) {
                const result = eventResult(event);
                if (result && event.kind === 'macro-finished') {
                    resultsByName.set(event.qualifiedName, result);
                }
            }
            const oracleIssues = validateVbaTestHostOracleTrace(events);
            for (const issue of oracleIssues) {
                log(`[runVbaTests oracle] ${issue.code}: ${issue.message}`);
            }
            resolve({
                events,
                resultsByName,
                hostError: hostError === undefined ? undefined : vbaTestFailureMessage(hostError),
                timedOutAfter,
            });
        };

        const killOwnedExcel = (reason: OwnedExcelKillReason) => {
            if (!ownedExcelPid || ownedExcelKilled) {
                return;
            }
            ownedExcelKilled = true;
            log(`[runVbaTests] Killing owned Excel process ${ownedExcelPid} after ${reason}.`);
            cp.spawn('taskkill.exe', ['/PID', String(ownedExcelPid), '/T', '/F']);
            const excelId = currentMacro?.excelId ?? events.find((event) => event.kind === 'excel-created')?.excelId;
            if (excelId) {
                events.push({ kind: 'excel-killed', excelId, reason });
            }
        };

        const armCleanupWatchdog = (stage: 'workbook-closed' | 'excel-quit') => {
            clearCleanupTimer();
            cleanupTimer = setTimeout(() => {
                if (settled) {
                    return;
                }
                const elapsedDescription = `${DEFAULT_VBA_TEST_CLEANUP_GRACE_MS} ms`;
                if (!sawExcelQuit) {
                    log(`[runVbaTests] Cleanup watchdog elapsed ${elapsedDescription} after workbook close; killing owned Excel.`);
                    killOwnedExcel('cleanup-failed');
                } else {
                    log(`[runVbaTests] Cleanup watchdog elapsed ${elapsedDescription} after Excel quit; stopping host script.`);
                }
                hostRun.kill();
                finish();
            }, DEFAULT_VBA_TEST_CLEANUP_GRACE_MS);
            log(`[runVbaTests] Cleanup watchdog armed after ${stage} (${DEFAULT_VBA_TEST_CLEANUP_GRACE_MS} ms).`);
        };

        startupTimer = setTimeout(() => {
            if (settled || currentMacro) {
                return;
            }
            timedOutAfter = 'test host startup';
            const excelId = events.find((event) => event.kind === 'excel-created')?.excelId ?? 'unknown';
            const message = `Excel test host timed out before starting tests after ${DEFAULT_VBA_TEST_TIMEOUT_MS} ms.`;
            events.push({
                kind: 'macro-finished',
                excelId,
                qualifiedName: 'XLIDE.TestHostStartup',
                outcome: 'timeout',
                durationMs: DEFAULT_VBA_TEST_TIMEOUT_MS,
                message,
            });
            killOwnedExcel('timeout');
            hostRun.kill();
            finish(message);
        }, DEFAULT_VBA_TEST_TIMEOUT_MS);

        const armMacroTimeout = (event: Extract<VbaTestHostOracleEvent, { kind: 'macro-started' }>) => {
            clearStartupTimer();
            clearCurrentTimer();
            const timeoutMs = event.timeoutMs ?? DEFAULT_VBA_TEST_TIMEOUT_MS;
            currentMacro = {
                excelId: event.excelId,
                qualifiedName: event.qualifiedName,
                timeoutMs,
                startedMs: Date.now(),
            };
            currentTimer = setTimeout(() => {
                if (!currentMacro || timedOutAfter) {
                    return;
                }
                timedOutAfter = currentMacro.qualifiedName;
                const durationMs = Date.now() - currentMacro.startedMs;
                const modalBlocker = currentModalBlocker;
                const modalDetail = modalBlocker
                    ? [modalBlocker.title, modalBlocker.message].filter(Boolean).join(': ')
                    : '';
                const outcome = modalBlocker ? 'modal-blocked' : 'timeout';
                const message = modalBlocker
                    ? `Blocked by Excel modal dialog${modalDetail ? ` (${modalDetail})` : ''}.`
                    : `Timed out after ${currentMacro.timeoutMs} ms.`;
                events.push({
                    kind: 'macro-finished',
                    excelId: currentMacro.excelId,
                    qualifiedName: currentMacro.qualifiedName,
                    outcome,
                    durationMs,
                    message,
                });
                killOwnedExcel(modalBlocker ? 'modal-blocked' : 'timeout');
                hostRun.kill();
                finish(message);
            }, timeoutMs);
        };

        const modalBlockedMessage = (event: Extract<VbaTestHostOracleEvent, { kind: 'modal-blocked' }>): string => {
            const modalDetail = [event.title, event.message].filter(Boolean).join(': ');
            return `Blocked by Excel modal dialog${modalDetail ? ` (${modalDetail})` : ''}.`;
        };

        const handleEvent = (event: VbaTestHostOracleEvent) => {
            events.push(event);
            if (event.kind === 'excel-created') {
                ownedExcelPid = event.pid;
                log(`[runVbaTests host] excel-created pid=${ownedExcelPid ?? 'unknown'}`);
            } else if (event.kind === 'host-phase') {
                log(`[runVbaTests host] phase ${event.phase} ${event.outcome} (${event.durationMs} ms)`);
            } else if (event.kind === 'macro-started') {
                log(`[runVbaTests host] macro-started ${event.qualifiedName} timeoutMs=${event.timeoutMs ?? DEFAULT_VBA_TEST_TIMEOUT_MS}`);
                currentModalBlocker = undefined;
                armMacroTimeout(event);
            } else if (event.kind === 'modal-detected') {
                log(`[runVbaTests host] modal-detected ${event.qualifiedName} classification=${event.classification ?? 'unknown'} safeToDismiss=${event.safeToDismiss ?? false}`);
            } else if (event.kind === 'modal-dismissed') {
                log(`[runVbaTests host] modal-dismissed ${event.qualifiedName} button=${event.button ?? 'unknown'} dismissed=${event.dismissed}`);
            } else if (event.kind === 'modal-blocked') {
                if (currentMacro &&
                    currentMacro.excelId === event.excelId &&
                    currentMacro.qualifiedName === event.qualifiedName) {
                    const macro = currentMacro;
                    currentModalBlocker = event;
                    const durationMs = Date.now() - macro.startedMs;
                    const message = modalBlockedMessage(event);
                    events.push({
                        kind: 'macro-finished',
                        excelId: macro.excelId,
                        qualifiedName: macro.qualifiedName,
                        outcome: 'modal-blocked',
                        durationMs,
                        message,
                    });
                    killOwnedExcel('modal-blocked');
                    hostRun.kill();
                    finish(message);
                }
                log(`[runVbaTests host] modal-blocked ${event.qualifiedName}: ${event.reason}`);
            } else if (event.kind === 'macro-finished') {
                log(`[runVbaTests host] macro-finished ${event.qualifiedName} outcome=${event.outcome}`);
                clearCurrentTimer();
                currentMacro = undefined;
                currentModalBlocker = undefined;
            } else if (event.kind === 'workbook-closed') {
                sawWorkbookClosed = true;
                log(`[runVbaTests host] workbook-closed durationMs=${event.durationMs ?? 'unknown'}`);
                armCleanupWatchdog('workbook-closed');
            } else if (event.kind === 'excel-quit') {
                sawExcelQuit = true;
                log(`[runVbaTests host] excel-quit durationMs=${event.durationMs ?? 'unknown'}`);
                if (sawWorkbookClosed) {
                    armCleanupWatchdog('excel-quit');
                }
            }
        };

        const handleStdoutLine = (line: string) => {
            const trimmed = line.trim();
            if (!trimmed) {
                return;
            }
            try {
                const event = parseVbaTestHostEventLine(trimmed);
                if (event) {
                    handleEvent(event);
                } else {
                    log(`[runVbaTests host stdout] ${trimmed}`);
                }
            } catch (err) {
                log(`[runVbaTests host stdout] ${trimmed}`);
                log(`[runVbaTests host parse] ${errorMessage(err)}`);
            }
        };

        const hostRun = runPowerShell({
            args: ['-File', hostScriptPath],
            onSpawn: (pid) => {
                log(`[runVbaTests] Spawned owned host powershell.exe (pid=${pid ?? 'unknown'})`);
            },
            onStdoutLine: handleStdoutLine,
            onStderrLine: (line) => {
                log(`[runVbaTests host stderr] ${line}`);
            },
        });
        void hostRun.result.then((result) => {
            if (result.spawnError) {
                const message = `RUNNER_FAILED|${result.spawnError.message}`;
                if (currentMacro) {
                    events.push({
                        kind: 'macro-finished',
                        excelId: currentMacro.excelId,
                        qualifiedName: currentMacro.qualifiedName,
                        outcome: 'runner-error',
                        durationMs: Date.now() - currentMacro.startedMs,
                        message,
                    });
                }
                killOwnedExcel('runner-error');
                finish(message);
                return;
            }
            log(`[runVbaTests] owned host powershell exited with code=${result.code} signal=${result.signal ?? 'none'}`);
            if (settled) {
                return;
            }
            if (result.code === 0) {
                finish();
                return;
            }
            const sentinel = result.stderrLines.find((line) => line.includes('XLIDE_TEST_HOST_ERROR|'));
            const hostError = sentinel
                ? sentinel.slice(sentinel.indexOf('XLIDE_TEST_HOST_ERROR|') + 'XLIDE_TEST_HOST_ERROR|'.length)
                : result.stderrLines.join('\n') || `PowerShell exited with code ${result.code}`;
            if (currentMacro) {
                events.push({
                    kind: 'macro-finished',
                    excelId: currentMacro.excelId,
                    qualifiedName: currentMacro.qualifiedName,
                    outcome: 'runner-error',
                    durationMs: Date.now() - currentMacro.startedMs,
                    message: hostError,
                });
            }
            killOwnedExcel('runner-error');
            finish(hostError);
        });
    });
    });
}

export function vbaTestRunItemFromHostResult(
    test: VbaTestCase,
    hostResult: OwnedReadOnlyExcelHostTestResult,
): VbaTestRunItem {
    const message = hostResult.message ? vbaTestFailureMessage(hostResult.message) : undefined;
    const expectedError = test.metadata.expectedError;
    if (expectedError !== undefined && (hostResult.outcome === 'passed' || hostResult.outcome === 'failed')) {
        return vbaTestRunItemFromExpectedError(test, hostResult, expectedError, message);
    }
    if (hostResult.outcome === 'passed') {
        return passedVbaTestRunItem(test, hostResult.durationMs, hostResult.output);
    }
    if (hostResult.outcome === 'failed') {
        return failedVbaTestRunItem(test, hostResult.durationMs, message, hostResult.output);
    }
    if (hostResult.outcome === 'timeout') {
        return {
            test,
            status: 'timeout',
            durationMs: hostResult.durationMs,
            error: message ?? `Timed out after ${hostResult.durationMs} ms.`,
            ...(hostResult.output?.length ? { output: hostResult.output } : {}),
        };
    }
    if (hostResult.outcome === 'modal-blocked') {
        return {
            test,
            status: 'host-error',
            durationMs: hostResult.durationMs,
            error: message ?? 'Blocked by an Excel modal dialog.',
            ...(hostResult.output?.length ? { output: hostResult.output } : {}),
        };
    }
    return {
        test,
        status: 'host-error',
        durationMs: hostResult.durationMs,
        error: message ?? 'The Excel test host failed while running this test.',
        ...(hostResult.output?.length ? { output: hostResult.output } : {}),
    };
}

function vbaTestRunItemFromExpectedError(
    test: VbaTestCase,
    hostResult: OwnedReadOnlyExcelHostTestResult,
    expectedError: number | 'any',
    failureMessage: string | undefined,
): VbaTestRunItem {
    if (hostResult.outcome === 'passed') {
        return failedVbaTestRunItem(
            test,
            hostResult.durationMs,
            expectedError === 'any'
                ? 'Expected a VBA error, but no error was raised.'
                : `Expected VBA error ${expectedError}, but no error was raised.`,
            hostResult.output,
        );
    }
    if (expectedError === 'any') {
        return passedVbaTestRunItem(test, hostResult.durationMs, hostResult.output);
    }
    if (hostResult.errorNumber === expectedError) {
        return passedVbaTestRunItem(test, hostResult.durationMs, hostResult.output);
    }
    const actual = hostResult.errorNumber !== undefined
        ? `VBA error ${hostResult.errorNumber}`
        : 'a failure without a deterministic VBA error number';
    return failedVbaTestRunItem(
        test,
        hostResult.durationMs,
        `Expected VBA error ${expectedError}, but got ${actual}${failureMessage ? `: ${failureMessage}` : '.'}`,
        hostResult.output,
    );
}

function passedVbaTestRunItem(test: VbaTestCase, durationMs: number, output?: string[]): VbaTestRunItem {
    if (test.metadata.xfailReason) {
        return {
            test,
            status: 'xpass',
            durationMs,
            error: `Expected failure did not occur: ${test.metadata.xfailReason}`,
            ...(output?.length ? { output } : {}),
        };
    }
    return { test, status: 'passed', durationMs, ...(output?.length ? { output } : {}) };
}

function failedVbaTestRunItem(
    test: VbaTestCase,
    durationMs: number,
    error: string | undefined,
    output?: string[],
): VbaTestRunItem {
    if (test.metadata.xfailReason) {
        return { test, status: 'xfail', durationMs, error, ...(output?.length ? { output } : {}) };
    }
    return { test, status: 'failed', durationMs, error, ...(output?.length ? { output } : {}) };
}

function logVbaTestRunItem(result: VbaTestRunItem, log: (message: string) => void): void {
    const detail = result.error ? `: ${result.error}` : '';
    log(`[runVbaTests] ${result.status.toUpperCase()} ${result.test.qualifiedName} (${result.durationMs} ms)${detail}`);
    for (const line of result.output ?? []) {
        log(`[runVbaTests output] ${result.test.qualifiedName}: ${line}`);
    }
}

function shouldFailFastStop(result: VbaTestRunItem, failFast?: boolean): boolean {
    return Boolean(failFast) &&
        (result.status === 'failed' ||
            result.status === 'xpass' ||
            result.status === 'timeout' ||
            result.status === 'host-error');
}
