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
export const DEFAULT_VBA_TEST_ARTIFACT_RETENTION = 20;
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
    host: VbaTestCiHostMetadata;
    durationMs: number;
}

export interface VbaTestCiHostMetadata {
    eventCount: number;
    excel: {
        created: number;
        quitNormally: boolean;
        killed: number;
        killReasons: string[];
    };
    modals: {
        detected: number;
        dismissed: number;
        blocked: number;
        blockedDialogs: VbaTestCiBlockedModal[];
    };
    phases: VbaTestCiHostPhaseSummary[];
}

export interface VbaTestCiBlockedModal {
    qualifiedName: string;
    title?: string;
    message?: string;
    buttons?: string[];
    buttonIds?: number[];
    reason: string;
}

export interface VbaTestCiHostPhaseSummary {
    phase: string;
    count: number;
    failed: number;
    totalDurationMs: number;
    maxDurationMs: number;
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
    retention?: number;
    generatedAt?: Date;
}

export interface VbaTestCiStatusOptions extends Pick<VbaTestRunArtifactOptions, 'generatedAt'> {
    hostEvents?: readonly VbaTestHostOracleEvent[];
}

export interface VbaTestRunArtifactWriteResult extends VbaTestRunArtifactPaths {
    ciStatus: VbaTestCiStatus;
}

const CI_FAILURE_STATUSES = new Set<VbaTestStatus>(['failed', 'timeout', 'host-error', 'xpass']);
const MAX_CI_MESSAGE_LENGTH = 500;

export function buildVbaTestRunArtifactPaths(
    report: Pick<VbaTestRunReport, 'filePath' | 'projectName' | 'startedAt'>,
    options: VbaTestRunArtifactOptions = {},
): VbaTestRunArtifactPaths {
    const outputFolder = resolveVbaTestArtifactOutputFolder(report.filePath, options.outputFolder);
    const runId = `${sanitizePathPart(path.basename(report.projectName, path.extname(report.projectName)))}_${formatRunTimestamp(new Date(report.startedAt))}`;
    const runDirectory = path.join(outputFolder, runId);
    const summaryPath = path.join(runDirectory, 'summary.json');
    const hostTracePath = path.join(runDirectory, 'host-trace.json');
    const outputLogPath = path.join(runDirectory, 'output.log');
    const statusPath = path.join(outputFolder, VBA_TEST_CI_STATUS_FILE_NAME);
    const projectDirectory = path.dirname(report.filePath);

    return {
        outputFolder,
        runDirectory,
        summaryPath,
        hostTracePath,
        outputLogPath,
        statusPath,
        runId,
        relativePaths: {
            runDirectory: relativeArtifactPath(projectDirectory, runDirectory),
            summary: relativeArtifactPath(projectDirectory, summaryPath),
            hostTrace: relativeArtifactPath(projectDirectory, hostTracePath),
            outputLog: relativeArtifactPath(projectDirectory, outputLogPath),
        },
    };
}

export function createVbaTestCiStatus(
    report: VbaTestRunReport,
    paths: VbaTestRunArtifactPaths,
    options: VbaTestCiStatusOptions = {},
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
            name: report.projectName,
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
        host: summarizeHostEventsForCi(options.hostEvents ?? []),
        durationMs: report.durationMs,
    };
}

export function sanitizeVbaTestHostTraceForArtifacts(
    events: readonly VbaTestHostOracleEvent[],
    projectPath: string,
): VbaTestHostOracleEvent[] {
    return events.map((event) => {
        if ('filePath' in event && typeof event.filePath === 'string') {
            return {
                ...event,
                filePath: artifactSafeProjectPath(event.filePath, projectPath),
            };
        }
        return { ...event };
    }) as VbaTestHostOracleEvent[];
}

export function renderVbaTestOutputLog(report: VbaTestRunReport, ciStatus: VbaTestCiStatus): string {
    const lines = [
        'XLIDE VBA Test Run',
        `Workbook: ${report.projectName}`,
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
    const sanitizedEvents = sanitizeVbaTestHostTraceForArtifacts(hostEvents, report.filePath);
    const ciStatus = createVbaTestCiStatus(report, paths, {
        ...options,
        hostEvents: sanitizedEvents,
    });
    const sanitizedHostTrace = {
        schemaVersion: 1,
        events: sanitizedEvents,
    };

    await fs.promises.mkdir(paths.runDirectory, { recursive: true });
    await fs.promises.writeFile(paths.summaryPath, jsonText(report), 'utf8');
    await fs.promises.writeFile(paths.hostTracePath, jsonText(sanitizedHostTrace), 'utf8');
    await fs.promises.writeFile(paths.outputLogPath, renderVbaTestOutputLog(report, ciStatus), 'utf8');
    await fs.promises.writeFile(paths.statusPath, jsonText(ciStatus), 'utf8');
    await pruneOldVbaTestRunArtifacts(paths, effectiveRetention(options.retention));

    return {
        ...paths,
        ciStatus,
    };
}

function resolveVbaTestArtifactOutputFolder(projectPath: string, configuredFolder?: string): string {
    const folder = configuredFolder?.trim() || DEFAULT_VBA_TEST_ARTIFACT_FOLDER;
    return path.isAbsolute(folder)
        ? path.normalize(folder)
        : path.join(path.dirname(projectPath), folder);
}

async function pruneOldVbaTestRunArtifacts(
    paths: VbaTestRunArtifactPaths,
    retention: number,
): Promise<void> {
    const entries = await xlideRunDirectoriesForWorkbook(paths.outputFolder, paths.runId);
    if (entries.length <= retention) {
        return;
    }

    const keep = new Set(entries
        .slice(0, retention)
        .map((entry) => entry.name));
    keep.add(paths.runId);

    // Retention pruning is best-effort: a locked old run directory (AV scan,
    // lingering handle) must not fail the current run's already-written
    // artifacts. Swallow individual rm failures rather than rejecting.
    await Promise.allSettled(entries
        .filter((entry) => !keep.has(entry.name))
        .map((entry) => fs.promises.rm(entry.fullPath, { recursive: true, force: true })));
}

async function xlideRunDirectoriesForWorkbook(
    outputFolder: string,
    currentRunId: string,
): Promise<Array<{ name: string; fullPath: string }>> {
    let entries: fs.Dirent[];
    try {
        entries = await fs.promises.readdir(outputFolder, { withFileTypes: true });
    } catch (err) {
        if (isNodeError(err) && err.code === 'ENOENT') {
            return [];
        }
        throw err;
    }

    const prefix = workbookRunIdPrefix(currentRunId);
    const candidates = entries
        .filter((entry) => entry.isDirectory() && isWorkbookRunDirectoryName(entry.name, prefix))
        .map((entry) => ({
            name: entry.name,
            fullPath: path.join(outputFolder, entry.name),
        }));
    const runDirectories: Array<{ name: string; fullPath: string }> = [];
    for (const candidate of candidates) {
        if (await isXlideRunArtifactDirectory(candidate.fullPath)) {
            runDirectories.push(candidate);
        }
    }
    return runDirectories.sort((left, right) => right.name.localeCompare(left.name));
}

function workbookRunIdPrefix(runId: string): string {
    const match = /^(.+_)\d{4}-\d{2}-\d{2}_\d{6}$/.exec(runId);
    return match?.[1] ?? '';
}

function isWorkbookRunDirectoryName(name: string, prefix: string): boolean {
    return prefix.length > 0 &&
        name.startsWith(prefix) &&
        /^\d{4}-\d{2}-\d{2}_\d{6}$/.test(name.slice(prefix.length));
}

function effectiveRetention(retention: number | undefined): number {
    return retention && Number.isInteger(retention) && retention > 0
        ? retention
        : DEFAULT_VBA_TEST_ARTIFACT_RETENTION;
}

async function isXlideRunArtifactDirectory(directory: string): Promise<boolean> {
    try {
        const summary = await fs.promises.stat(path.join(directory, 'summary.json'));
        return summary.isFile();
    } catch (err) {
        if (isNodeError(err) && err.code === 'ENOENT') {
            return false;
        }
        throw err;
    }
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

function summarizeHostEventsForCi(events: readonly VbaTestHostOracleEvent[]): VbaTestCiHostMetadata {
    const killReasons = events
        .filter((event): event is Extract<VbaTestHostOracleEvent, { kind: 'excel-killed' }> => event.kind === 'excel-killed')
        .map((event) => event.reason);
    const blockedDialogs = events
        .filter((event): event is Extract<VbaTestHostOracleEvent, { kind: 'modal-blocked' }> => event.kind === 'modal-blocked')
        .map((event) => ({
            qualifiedName: event.qualifiedName,
            ...(sanitizeCiMessage(event.title) ? { title: sanitizeCiMessage(event.title) } : {}),
            ...(sanitizeCiMessage(event.message) ? { message: sanitizeCiMessage(event.message) } : {}),
            ...(event.buttons?.length ? { buttons: event.buttons.map((button) => sanitizeCiMessage(button) ?? '') } : {}),
            ...(event.buttonIds?.length ? { buttonIds: event.buttonIds } : {}),
            reason: sanitizeCiMessage(event.reason) ?? 'unknown',
        }));
    return {
        eventCount: events.length,
        excel: {
            created: events.filter((event) => event.kind === 'excel-created').length,
            quitNormally: events.some((event) => event.kind === 'excel-quit'),
            killed: killReasons.length,
            killReasons,
        },
        modals: {
            detected: events.filter((event) => event.kind === 'modal-detected').length,
            dismissed: events.filter((event) => event.kind === 'modal-dismissed').length,
            blocked: blockedDialogs.length,
            blockedDialogs,
        },
        phases: summarizeHostPhaseDurations(events),
    };
}

function summarizeHostPhaseDurations(events: readonly VbaTestHostOracleEvent[]): VbaTestCiHostPhaseSummary[] {
    const summaries = new Map<string, VbaTestCiHostPhaseSummary>();
    for (const event of events) {
        if (event.kind !== 'host-phase') {
            continue;
        }
        const current = summaries.get(event.phase) ?? {
            phase: event.phase,
            count: 0,
            failed: 0,
            totalDurationMs: 0,
            maxDurationMs: 0,
        };
        current.count += 1;
        current.failed += event.outcome === 'failed' ? 1 : 0;
        current.totalDurationMs += event.durationMs;
        current.maxDurationMs = Math.max(current.maxDurationMs, event.durationMs);
        summaries.set(event.phase, current);
    }
    return [...summaries.values()].sort((left, right) => left.phase.localeCompare(right.phase));
}

function artifactSafeProjectPath(value: string, projectPath: string): string {
    const projectDirectory = path.dirname(projectPath);
    const relative = path.relative(projectDirectory, value);
    if (relative && !relative.startsWith('..') && !path.isAbsolute(relative)) {
        return toPosixPath(relative);
    }
    return path.basename(value);
}

function formatOutputLogResult(result: VbaTestRunItem): string {
    const detail = result.error ? ` - ${sanitizeCiMessage(result.error) ?? ''}` : '';
    const output = result.output?.length
        ? `\n${result.output.map((line) => `  output: ${sanitizeCiMessage(line) ?? ''}`).join('\n')}`
        : '';
    return `- ${result.status.toUpperCase()} ${result.test.qualifiedName} (${result.durationMs} ms)${detail}${output}`;
}

function formatRunTimestamp(date: Date): string {
    const iso = date.toISOString();
    return `${iso.slice(0, 10)}_${iso.slice(11, 19).replace(/:/g, '')}`;
}

function sanitizePathPart(value: string): string {
    const sanitized = value.trim().replace(/[^A-Za-z0-9._-]+/g, '_').replace(/^_+|_+$/g, '');
    return sanitized || 'workbook';
}

function relativeArtifactPath(projectDirectory: string, targetPath: string): string {
    return toPosixPath(path.relative(projectDirectory, targetPath) || '.');
}

function toPosixPath(value: string): string {
    return value.replace(/\\/g, '/');
}

function jsonText(value: unknown): string {
    return `${JSON.stringify(value, null, 2)}\n`;
}

function isNodeError(value: unknown): value is NodeJS.ErrnoException {
    return value !== null && typeof value === 'object' && 'code' in value;
}
