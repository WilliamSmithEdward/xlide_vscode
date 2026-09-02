import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

vi.mock('../src/vbaTestSupportStatus', () => ({ getVbaTestSupportStatus: vi.fn() }));
vi.mock('../src/excelComAvailability', () => ({ checkExcelComAvailability: vi.fn() }));
vi.mock('../src/vbaTestExecution', () => ({ runWorkbookVbaTests: vi.fn() }));
vi.mock('../src/vbaTestArtifacts', async (importOriginal) => {
    const actual = await importOriginal<typeof import('../src/vbaTestArtifacts')>();
    return { ...actual, writeVbaTestRunArtifacts: vi.fn(actual.writeVbaTestRunArtifacts) };
});

import type { ProjectEngine } from '../src/projectEngine';
import { checkExcelComAvailability, type ExcelComAvailabilityStatus } from '../src/excelComAvailability';
import { getVbaTestSupportStatus, type VbaTestSupportStatus } from '../src/vbaTestSupportStatus';
import { runWorkbookVbaTests, type VbaTestRunExecution } from '../src/vbaTestExecution';
import { writeVbaTestRunArtifacts } from '../src/vbaTestArtifacts';
import { executeVbaTestRun, type VbaTestRunPipelineRunner } from '../src/vbaTestRunPipeline';
import type { VbaTestCase, VbaTestRunReport } from '../src/vbaTestRunner';
import { writeProjectSettings } from '../src/projectSettings';

const tempRoots: string[] = [];

beforeEach(() => {
    vi.clearAllMocks();
});

afterEach(() => {
    for (const root of tempRoots.splice(0)) {
        fs.rmSync(root, { recursive: true, force: true });
    }
});

function bridge(): ProjectEngine {
    return {} as ProjectEngine;
}

function tempWorkbook(): string {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-test-run-pipeline-'));
    tempRoots.push(root);
    const workbook = path.join(root, 'Book.xlsm');
    fs.writeFileSync(workbook, '', 'utf8');
    return workbook;
}

function supportStatus(canRun: boolean): VbaTestSupportStatus {
    return {
        state: canRun ? 'installed' : 'missing',
        title: canRun ? 'XlideAssert.bas Installed' : 'XlideAssert.bas Not Installed',
        description: 'support status',
        actionLabel: canRun ? 'Installed' : 'Install',
        canInstall: !canRun,
        canRun,
    };
}

function comStatus(canRun: boolean): ExcelComAvailabilityStatus {
    return {
        state: canRun ? 'installed' : 'missing',
        title: canRun ? 'Excel COM Ready' : 'Excel COM Not Found',
        description: 'runtime status',
        canRun,
    };
}

function executionFor(workbook: string): VbaTestRunExecution {
    const test: VbaTestCase = {
        id: 'Tests.Test_Pass',
        moduleName: 'Tests',
        moduleType: 'standard',
        procedureName: 'Test_Pass',
        qualifiedName: 'Tests.Test_Pass',
        line: 4,
        column: 1,
        annotationLine: 3,
        metadata: { tags: [] },
    };
    const report: VbaTestRunReport = {
        filePath: workbook,
        projectName: path.basename(workbook),
        startedAt: '2026-06-03T12:34:56.000Z',
        durationMs: 42,
        discovery: {
            filePath: workbook,
            tests: [test],
            unfilteredTestCount: 1,
            modulesScanned: 1,
            modulesIgnored: 0,
            contract: 'contract',
        },
        results: [{ test, status: 'passed', durationMs: 5 }],
    };
    return { report, hostEvents: [] };
}

describe('executeVbaTestRun', () => {
    it('blocks on missing test support before probing Excel COM or running tests', async () => {
        vi.mocked(getVbaTestSupportStatus).mockResolvedValue(supportStatus(false));

        const result = await executeVbaTestRun(bridge(), 'C:/work/Book.xlsm');

        expect(result).toEqual({ kind: 'blocked-support', support: supportStatus(false) });
        expect(checkExcelComAvailability).not.toHaveBeenCalled();
        expect(runWorkbookVbaTests).not.toHaveBeenCalled();
    });

    it('blocks on unavailable Excel COM without running tests', async () => {
        vi.mocked(getVbaTestSupportStatus).mockResolvedValue(supportStatus(true));
        vi.mocked(checkExcelComAvailability).mockResolvedValue(comStatus(false));

        const result = await executeVbaTestRun(bridge(), 'C:/work/Book.xlsm');

        expect(result).toEqual({ kind: 'blocked-com', runtime: comStatus(false) });
        expect(runWorkbookVbaTests).not.toHaveBeenCalled();
    });

    it('runs tests through the caller wrapper and writes artifacts once', async () => {
        const workbook = tempWorkbook();
        await writeProjectSettings(workbook, {
            tests: {
                artifactFolder: 'ci-tests',
                artifactRetention: 3,
            },
        });
        vi.mocked(getVbaTestSupportStatus).mockResolvedValue(supportStatus(true));
        vi.mocked(checkExcelComAvailability).mockResolvedValue(comStatus(true));
        const execution = executionFor(workbook);
        vi.mocked(runWorkbookVbaTests).mockResolvedValue(execution);
        const progress = { report: vi.fn() };
        const log = vi.fn();
        const runTests = vi.fn((run: VbaTestRunPipelineRunner) => run(progress));

        const result = await executeVbaTestRun(bridge(), workbook, {
            selection: { moduleName: 'Tests' },
            failFast: true,
            log,
            runTests,
        });

        expect(runTests).toHaveBeenCalledTimes(1);
        expect(runWorkbookVbaTests).toHaveBeenCalledWith(expect.anything(), workbook, {
            selection: { moduleName: 'Tests' },
            failFast: true,
            log,
            progress,
        });
        expect(result.kind).toBe('completed');
        if (result.kind !== 'completed') {
            throw new Error('expected a completed pipeline result');
        }
        expect(result.execution).toBe(execution);
        expect(result.artifacts.ok).toBe(true);
        if (!result.artifacts.ok) {
            throw new Error(result.artifacts.error);
        }
        expect(result.artifacts.settings).toMatchObject({
            artifactFolder: 'ci-tests',
            artifactFolderSource: 'project',
            artifactRetention: 3,
            artifactRetentionSource: 'project',
        });
        expect(fs.existsSync(result.artifacts.artifacts.statusPath)).toBe(true);
    });

    it('refuses a second run while one is in flight, then allows the next', async () => {
        const workbook = tempWorkbook();
        vi.mocked(getVbaTestSupportStatus).mockResolvedValue(supportStatus(true));
        vi.mocked(checkExcelComAvailability).mockResolvedValue(comStatus(true));
        let releaseFirst!: (execution: VbaTestRunExecution) => void;
        vi.mocked(runWorkbookVbaTests).mockReturnValueOnce(
            new Promise<VbaTestRunExecution>((resolve) => { releaseFirst = resolve; }),
        );

        const first = executeVbaTestRun(bridge(), workbook);
        // The guard takes effect synchronously at entry, so even a run still
        // inside its support/COM gates blocks the next one - including runs
        // for a different file: single-instance hosts cannot share.
        const second = await executeVbaTestRun(bridge(), 'C:/other/Other.xlsm');

        expect(second).toEqual({
            kind: 'blocked-busy',
            activeRunDescription: path.basename(workbook),
        });

        releaseFirst(executionFor(workbook));
        const firstResult = await first;
        expect(firstResult.kind).toBe('completed');

        vi.mocked(runWorkbookVbaTests).mockResolvedValue(executionFor(workbook));
        const third = await executeVbaTestRun(bridge(), workbook);
        expect(third.kind).toBe('completed');
    });

    it('clears the busy guard when a run throws', async () => {
        const workbook = tempWorkbook();
        vi.mocked(getVbaTestSupportStatus).mockResolvedValue(supportStatus(true));
        vi.mocked(checkExcelComAvailability).mockResolvedValue(comStatus(true));
        vi.mocked(runWorkbookVbaTests).mockRejectedValueOnce(new Error('host died'));

        await expect(executeVbaTestRun(bridge(), workbook)).rejects.toThrow('host died');

        vi.mocked(runWorkbookVbaTests).mockResolvedValue(executionFor(workbook));
        const next = await executeVbaTestRun(bridge(), workbook);
        expect(next.kind).toBe('completed');
    });

    it('reports artifact write failures without failing the run', async () => {
        const workbook = tempWorkbook();
        vi.mocked(getVbaTestSupportStatus).mockResolvedValue(supportStatus(true));
        vi.mocked(checkExcelComAvailability).mockResolvedValue(comStatus(true));
        vi.mocked(runWorkbookVbaTests).mockResolvedValue(executionFor(workbook));
        vi.mocked(writeVbaTestRunArtifacts).mockRejectedValueOnce(new Error('disk full'));

        const result = await executeVbaTestRun(bridge(), workbook);

        expect(result.kind).toBe('completed');
        if (result.kind !== 'completed') {
            throw new Error('expected a completed pipeline result');
        }
        expect(result.artifacts).toEqual({ ok: false, error: 'disk full' });
    });
});
