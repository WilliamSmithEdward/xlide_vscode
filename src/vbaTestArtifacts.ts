import * as fs from 'fs';
import * as path from 'path';
import {
    summarizeVbaTestRun,
    type VbaTestRunItem,
    type VbaTestRunReport,
    type VbaTestRunSummary,
    type VbaTestStatus,
} from './vbaTestRunner';
import type { VbaTestHostOracleEvent } from './vbaTestHostOracle';

export const DEFAULT_VBA_TEST_ARTIFACT_FOLDER = 'tests';
export const VBA_TEST_CI_STATUS_FILE_NAME = 'status_for_ci.json';

export type VbaTestCiStatusValue = 'pass' | 'fail' | 'error';
export type VbaTestCiReason =
    | 'passed'
    | 'test-failures'
    | 'timeouts'
    | 'host-errors'
    | 'unexpected-pass'
    | 'no-tests'
    | 'runner-error';

export interface VbaTestRunArtifactPaths {
    outputFolder: string;
    runDirectory: string;
    summaryPath: string;
    hostTracePath: string;
    outputLogPath: string;
    statusPath: string;
    runId: string;
    relativePaths: {
        runDirectory: string;
        summary: string;
        hostTrace: string;
        outputLog: string;
    };
}

export interface VbaTestCiStatus {
    schemaVersion: 1;
    status: VbaTestCiStatusValue;
    reason: VbaTestCiReason;
    generatedAt: string;
    runId: string;
    workbook: {
        name: string;
    };
    paths: {
        runDirectory: string;
        summary: string;
        hostTrace: string;
        outputLog: string;
    };
    counts: {
        total: number;
        passed: number;
        failed: number;
        timeout: number;
        hostError: number;
        skipped: number;
        xfail: number;
        xpass: number;
    };
    failedTests: VbaTestCiFailedTest[];
    durationMs: number;
}

export interface VbaTestCiFailedTest {
    id: string;
    module: string;
    procedure: string;
    status: Extract<VbaTestStatus, 'failed' | 'timeout' | 'host-error' | 'xpass'>;
    durationMs: number;
    message?: string;
}

export interface VbaTestRunArtifactOptions {
    outputFolder?: string;
    generatedAt?: Date;
}

export interface VbaTestRunArtifactWriteResult extends VbaTestRunArtifactPaths {
    ciStatus: VbaTestCiStatus;
}

const CI_FAILURE_STATUSES = new Set<VbaTestStatus>(['failed', 'timeout', 'host-error', 'xpass']);
const MAX_CI_MESSAGE_LENGTH = 500;

export function buildVbaTestRunArtifactPaths(
    report: Pick<VbaTestRunReport, 'filePath' | 'workbookName' | 'startedAt'>,
    options: VbaTestRunArtifactOptions = {},
): VbaTestRunArtifactPaths {
    const outputFolder = resolveVbaTestArtifactOutputFolder(report.filePath, options.outputFolder);
    const runId = `${sanitizePathPart(path.basename(report.workbookName, path.extname(report.workbookName)))}_${formatRunTimestamp(new Date(report.startedAt))}`;
    const runDirectory = path.join(outputFolder, runId);
    const summaryPath = path.join(runDirectory, 'summary.json');
    const hostTracePath = path.join(runDirectory, 'host-trace.json');
    const outputLogPath = path.join(runDirectory, 'output.log');
    const statusPath = path.join(outputFolder, VBA_TEST_CI_STATUS_FILE_NAME);
    const workbookDirectory = path.dirname(report.filePath);

    return {
        outputFolder,
        runDirectory,
        summaryPath,
        hostTracePath,
        outputLogPath,
        statusPath,
        runId,
        relativePaths: {
            runDirectory: relativeArtifactPath(workbookDirectory, runDirectory),
            summary: relativeArtifactPath(workbookDirectory, summaryPath),
            hostTrace: relativeArtifactPath(workbookDirectory, hostTracePath),
            outputLog: relativeArtifactPath(workbookDirectory, outputLogPath),
        },
    };
}

export function createVbaTestCiStatus(
    report: VbaTestRunReport,
    paths: VbaTestRunArtifactPaths,
    options: Pick<VbaTestRunArtifactOptions, 'generatedAt'> = {},
): VbaTestCiStatus {
    const summary = summarizeVbaTestRun(report);
    const classification = classifyCiStatus(summary);
    return {
        schemaVersion: 1,
        status: classification.status,
        reason: classification.reason,
        generatedAt: (options.generatedAt ?? new Date()).toISOString(),
        runId: paths.runId,
        workbook: {
            name: report.workbookName,
        },
        paths: {
            runDirectory: paths.relativePaths.runDirectory,
            summary: paths.relativePaths.summary,
            hostTrace: paths.relativePaths.hostTrace,
            outputLog: paths.relativePaths.outputLog,
        },
        counts: {
            total: summary.total,
            passed: summary.passed,
            failed: summary.failed,
            timeout: summary.timeout,
            hostError: summary.hostError,
            skipped: summary.skipped,
            xfail: summary.xfail,
            xpass: summary.xpass,
        },
        failedTests: report.results
            .filter((result) => CI_FAILURE_STATUSES.has(result.status))
            .map(ciFailedTest),
        durationMs: report.durationMs,
    };
}

export function sanitizeVbaTestHostTraceForArtifacts(
    events: readonly VbaTestHostOracleEvent[],
    workbookPath: string,
): VbaTestHostOracleEvent[] {
    return events.map((event) => {
        if ('filePath' in event && typeof event.filePath === 'string') {
            return {
                ...event,
                filePath: artifactSafeWorkbookPath(event.filePath, workbookPath),
            };
        }
        return { ...event };
    }) as VbaTestHostOracleEvent[];
}

export function renderVbaTestOutputLog(report: VbaTestRunReport, ciStatus: VbaTestCiStatus): string {
    const lines = [
        'XLIDE VBA Test Run',
        `Workbook: ${report.workbookName}`,
        `Started: ${report.startedAt}`,
        `Duration: ${report.durationMs} ms`,
        `Status: ${ciStatus.status} (${ciStatus.reason})`,
        `Counts: total=${ciStatus.counts.total} passed=${ciStatus.counts.passed} failed=${ciStatus.counts.failed} timeout=${ciStatus.counts.timeout} hostError=${ciStatus.counts.hostError} skipped=${ciStatus.counts.skipped} xfail=${ciStatus.counts.xfail} xpass=${ciStatus.counts.xpass}`,
        '',
        'Results:',
        ...report.results.map(formatOutputLogResult),
    ];
    return `${lines.join('\n')}\n`;
}

export async function writeVbaTestRunArtifacts(
    report: VbaTestRunReport,
    hostEvents: readonly VbaTestHostOracleEvent[] = [],
    options: VbaTestRunArtifactOptions = {},
): Promise<VbaTestRunArtifactWriteResult> {
    const paths = buildVbaTestRunArtifactPaths(report, options);
    const ciStatus = createVbaTestCiStatus(report, paths, options);
    const sanitizedHostTrace = {
        schemaVersion: 1,
        events: sanitizeVbaTestHostTraceForArtifacts(hostEvents, report.filePath),
    };

    await fs.promises.mkdir(paths.runDirectory, { recursive: true });
    await fs.promises.writeFile(paths.summaryPath, jsonText(report), 'utf8');
    await fs.promises.writeFile(paths.hostTracePath, jsonText(sanitizedHostTrace), 'utf8');
    await fs.promises.writeFile(paths.outputLogPath, renderVbaTestOutputLog(report, ciStatus), 'utf8');
    await fs.promises.writeFile(paths.statusPath, jsonText(ciStatus), 'utf8');

    return {
        ...paths,
        ciStatus,
    };
}

function resolveVbaTestArtifactOutputFolder(workbookPath: string, configuredFolder?: string): string {
    const folder = configuredFolder?.trim() || DEFAULT_VBA_TEST_ARTIFACT_FOLDER;
    return path.isAbsolute(folder)
        ? path.normalize(folder)
        : path.join(path.dirname(workbookPath), folder);
}

function classifyCiStatus(summary: VbaTestRunSummary): { status: VbaTestCiStatusValue; reason: VbaTestCiReason } {
    if (summary.total === 0) {
        return { status: 'error', reason: 'no-tests' };
    }
    if (summary.hostError > 0) {
        return { status: 'error', reason: 'host-errors' };
    }
    if (summary.timeout > 0) {
        return { status: 'fail', reason: 'timeouts' };
    }
    if (summary.failed > 0) {
        return { status: 'fail', reason: 'test-failures' };
    }
    if (summary.xpass > 0) {
        return { status: 'fail', reason: 'unexpected-pass' };
    }
    return { status: 'pass', reason: 'passed' };
}

function ciFailedTest(result: VbaTestRunItem): VbaTestCiFailedTest {
    const failed: VbaTestCiFailedTest = {
        id: result.test.id,
        module: result.test.moduleName,
        procedure: result.test.procedureName,
        status: result.status as VbaTestCiFailedTest['status'],
        durationMs: result.durationMs,
    };
    const message = sanitizeCiMessage(result.error);
    if (message) {
        failed.message = message;
    }
    return failed;
}

function sanitizeCiMessage(value: string | undefined): string | undefined {
    if (!value) {
        return undefined;
    }
    const sanitized = value
        .replace(/[\r\n]+/g, ' ')
        .replace(/[\u0000-\u001F\u007F]+/g, ' ')
        .replace(/ {2,}/g, ' ')
        .trim();
    if (sanitized.length <= MAX_CI_MESSAGE_LENGTH) {
        return sanitized;
    }
    return `${sanitized.slice(0, MAX_CI_MESSAGE_LENGTH - 3)}...`;
}

function artifactSafeWorkbookPath(value: string, workbookPath: string): string {
    const workbookDirectory = path.dirname(workbookPath);
    const relative = path.relative(workbookDirectory, value);
    if (relative && !relative.startsWith('..') && !path.isAbsolute(relative)) {
        return toPosixPath(relative);
    }
    return path.basename(value);
}

function formatOutputLogResult(result: VbaTestRunItem): string {
    const detail = result.error ? ` - ${sanitizeCiMessage(result.error) ?? ''}` : '';
    return `- ${result.status.toUpperCase()} ${result.test.qualifiedName} (${result.durationMs} ms)${detail}`;
}

function formatRunTimestamp(date: Date): string {
    const iso = date.toISOString();
    return `${iso.slice(0, 10)}_${iso.slice(11, 19).replace(/:/g, '')}`;
}

function sanitizePathPart(value: string): string {
    const sanitized = value.trim().replace(/[^A-Za-z0-9._-]+/g, '_').replace(/^_+|_+$/g, '');
    return sanitized || 'workbook';
}

function relativeArtifactPath(workbookDirectory: string, targetPath: string): string {
    return toPosixPath(path.relative(workbookDirectory, targetPath) || '.');
}

function toPosixPath(value: string): string {
    return value.replace(/\\/g, '/');
}

function jsonText(value: unknown): string {
    return `${JSON.stringify(value, null, 2)}\n`;
}
