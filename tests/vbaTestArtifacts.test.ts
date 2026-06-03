import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
    buildVbaTestRunArtifactPaths,
    createVbaTestCiStatus,
    sanitizeVbaTestHostTraceForArtifacts,
    writeVbaTestRunArtifacts,
} from '../src/vbaTestArtifacts';
import type { VbaTestHostOracleEvent } from '../src/vbaTestHostOracle';
import type { VbaTestCase, VbaTestRunItem, VbaTestRunReport, VbaTestStatus } from '../src/vbaTestRunner';

const tempRoots: string[] = [];

afterEach(() => {
    for (const root of tempRoots.splice(0)) {
        fs.rmSync(root, { recursive: true, force: true });
    }
});

describe('VBA test artifacts', () => {
    it('builds deterministic CI status without leaking paths or nondeterministic locations', () => {
        const workbook = tempWorkbook('Live Test.xlsm');
        const report = testReport(workbook, [
            result('Tests.Passes', 'passed', 10),
            result('Tests.Fails', 'failed', 11, 'expected 2\r\nactual 3'),
            result('Tests.Hangs', 'timeout', 30000, 'Timed out\u0007after 30000 ms'),
            result('Tests.HostBlocked', 'host-error', 0, 'Protected View blocked automation'),
            result('Tests.ExpectedButPassed', 'xpass', 2, 'Unexpected pass'),
            result('Tests.KnownBug', 'xfail', 3, 'Known failure'),
        ]);
        const paths = buildVbaTestRunArtifactPaths(report);

        expect(paths.runId).toBe('Live_Test_2026-06-03_212233');
        expect(paths.relativePaths).toEqual({
            runDirectory: 'tests/Live_Test_2026-06-03_212233',
            summary: 'tests/Live_Test_2026-06-03_212233/summary.json',
            hostTrace: 'tests/Live_Test_2026-06-03_212233/host-trace.json',
            outputLog: 'tests/Live_Test_2026-06-03_212233/output.log',
        });

        const status = createVbaTestCiStatus(report, paths, {
            generatedAt: new Date('2026-06-03T21:23:00.000Z'),
        });

        expect(status).toMatchObject({
            schemaVersion: 1,
            status: 'error',
            reason: 'host-errors',
            generatedAt: '2026-06-03T21:23:00.000Z',
            workbook: { name: 'Live Test.xlsm' },
            counts: {
                total: 6,
                passed: 1,
                failed: 1,
                timeout: 1,
                hostError: 1,
                skipped: 0,
                xfail: 1,
                xpass: 1,
            },
        });
        expect(status.failedTests).toEqual([
            {
                id: 'Tests.Fails',
                module: 'Tests',
                procedure: 'Fails',
                status: 'failed',
                durationMs: 11,
                message: 'expected 2 actual 3',
            },
            {
                id: 'Tests.Hangs',
                module: 'Tests',
                procedure: 'Hangs',
                status: 'timeout',
                durationMs: 30000,
                message: 'Timed out after 30000 ms',
            },
            {
                id: 'Tests.HostBlocked',
                module: 'Tests',
                procedure: 'HostBlocked',
                status: 'host-error',
                durationMs: 0,
                message: 'Protected View blocked automation',
            },
            {
                id: 'Tests.ExpectedButPassed',
                module: 'Tests',
                procedure: 'ExpectedButPassed',
                status: 'xpass',
                durationMs: 2,
                message: 'Unexpected pass',
            },
        ]);
        expect(JSON.stringify(status)).not.toContain(workbook);
        expect(status.failedTests[0]).not.toHaveProperty('line');
        expect(status.failedTests[0]).not.toHaveProperty('column');
    });

    it('treats zero discovered tests as an error for CI consumers', () => {
        const workbook = tempWorkbook('Empty.xlsm');
        const report = testReport(workbook, []);
        const paths = buildVbaTestRunArtifactPaths(report);

        expect(createVbaTestCiStatus(report, paths, {
            generatedAt: new Date('2026-06-03T21:23:00.000Z'),
        })).toMatchObject({
            status: 'error',
            reason: 'no-tests',
            counts: {
                total: 0,
            },
            failedTests: [],
        });
    });

    it('writes summary, sanitized host trace, output log, and latest CI status', async () => {
        const workbook = tempWorkbook('Live Test.xlsm');
        const report = testReport(workbook, [
            result('Tests.Passes', 'passed', 10),
        ]);
        const hostEvents: VbaTestHostOracleEvent[] = [
            { kind: 'excel-created', excelId: 'xlide-1', owned: true, pid: 123 },
            {
                kind: 'workbook-opened',
                excelId: 'xlide-1',
                filePath: workbook,
                readOnly: true,
                updateLinks: 0,
                displayAlerts: false,
                ignoreReadOnlyRecommended: true,
            },
            { kind: 'macro-started', excelId: 'xlide-1', qualifiedName: 'Tests.Passes', timeoutMs: 30000 },
            { kind: 'macro-finished', excelId: 'xlide-1', qualifiedName: 'Tests.Passes', outcome: 'passed', durationMs: 10 },
            { kind: 'workbook-closed', excelId: 'xlide-1', filePath: workbook, saveChanges: false },
            { kind: 'excel-quit', excelId: 'xlide-1' },
        ];

        const written = await writeVbaTestRunArtifacts(report, hostEvents, {
            generatedAt: new Date('2026-06-03T21:23:00.000Z'),
        });

        expect(fs.existsSync(written.summaryPath)).toBe(true);
        expect(fs.existsSync(written.hostTracePath)).toBe(true);
        expect(fs.existsSync(written.outputLogPath)).toBe(true);
        expect(fs.existsSync(written.statusPath)).toBe(true);

        expect(JSON.parse(fs.readFileSync(written.summaryPath, 'utf8'))).toMatchObject({
            workbookName: 'Live Test.xlsm',
            results: [{ status: 'passed' }],
        });
        const hostTrace = JSON.parse(fs.readFileSync(written.hostTracePath, 'utf8'));
        expect(hostTrace.schemaVersion).toBe(1);
        expect(hostTrace.events[0]).toMatchObject({ kind: 'excel-created' });
        expect(hostTrace.events[1]).toMatchObject({ kind: 'workbook-opened', filePath: 'Live Test.xlsm' });
        expect(hostTrace.events[4]).toMatchObject({ kind: 'workbook-closed', filePath: 'Live Test.xlsm' });
        expect(fs.readFileSync(written.outputLogPath, 'utf8')).toContain('Status: pass (passed)');
        expect(JSON.parse(fs.readFileSync(written.statusPath, 'utf8'))).toMatchObject({
            status: 'pass',
            reason: 'passed',
            paths: {
                summary: 'tests/Live_Test_2026-06-03_212233/summary.json',
            },
        });
    });

    it('sanitizes workbook paths in persisted host traces', () => {
        const workbook = tempWorkbook('Book.xlsm');
        const events: VbaTestHostOracleEvent[] = [
            {
                kind: 'workbook-opened',
                excelId: 'xlide-1',
                filePath: workbook,
                readOnly: true,
            },
            {
                kind: 'workbook-closed',
                excelId: 'xlide-1',
                filePath: path.join(path.dirname(workbook), 'Other.xlsm'),
                saveChanges: false,
            },
        ];

        expect(sanitizeVbaTestHostTraceForArtifacts(events, workbook)).toEqual([
            expect.objectContaining({ filePath: 'Book.xlsm' }),
            expect.objectContaining({ filePath: 'Other.xlsm' }),
        ]);
    });
});

function tempWorkbook(name: string): string {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-vba-test-artifacts-'));
    tempRoots.push(root);
    return path.join(root, name);
}

function testReport(workbook: string, results: VbaTestRunItem[]): VbaTestRunReport {
    const tests = results.map((entry) => entry.test);
    return {
        filePath: workbook,
        workbookName: path.basename(workbook),
        startedAt: '2026-06-03T21:22:33.000Z',
        durationMs: results.reduce((total, entry) => total + entry.durationMs, 0),
        discovery: {
            filePath: workbook,
            tests,
            unfilteredTestCount: tests.length,
            modulesScanned: 1,
            modulesIgnored: 0,
            contract: 'contract',
        },
        results,
    };
}

function result(qualifiedName: string, status: VbaTestStatus, durationMs: number, error?: string): VbaTestRunItem {
    return {
        test: testCase(qualifiedName),
        status,
        durationMs,
        ...(error ? { error } : {}),
    };
}

function testCase(qualifiedName: string): VbaTestCase {
    const [moduleName, procedureName] = qualifiedName.split('.');
    return {
        id: qualifiedName,
        moduleName,
        moduleType: 'standard',
        procedureName,
        qualifiedName,
        line: 2,
        column: 1,
        annotationLine: 1,
        metadata: { tags: [] },
    };
}
