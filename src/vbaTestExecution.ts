import * as path from 'path';
import type { PythonBridge } from './pythonBridge';
import type { VbaTestHostOracleEvent } from './vbaTestHostOracle';
import {
    runOwnedExcelTestHostSession,
    type OwnedReadOnlyExcelHostRunResult,
    type OwnedReadOnlyExcelHostTestResult,
} from './vbaTestHostSession';
import { stageOwnedReadOnlyExcelTestHost } from './vbaTestHostStaging';
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

export type { OwnedReadOnlyExcelHostTestResult } from './vbaTestHostSession';

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
    const staging = await stageOwnedReadOnlyExcelTestHost(bridge, filePath, tests, {
        failFast: options.failFast,
        log,
    });
    log(`[runVbaTests] Running owned read-only Excel host for ${tests.length} test(s).`);
    log(`[runVbaTests] Temporary workbook path: ${staging.tempWorkbookPath}`);
    log(`[runVbaTests] Host script path: ${staging.hostScriptPath}`);
    return runOwnedExcelTestHostSession({
        hostScriptPath: staging.hostScriptPath,
        disposeStaging: staging.dispose,
        log,
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
