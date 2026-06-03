import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { writeAgentVbaTestArtifacts } from '../src/agentVbaTestArtifacts';
import type { VbaTestCase, VbaTestRunReport } from '../src/vbaTestRunner';
import { writeWorkbookSettings } from '../src/workbookSettings';

const tempRoots: string[] = [];

afterEach(() => {
    for (const root of tempRoots.splice(0)) {
        fs.rmSync(root, { recursive: true, force: true });
    }
});

function tempWorkbook(): { root: string; workbook: string } {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-agent-test-artifacts-'));
    tempRoots.push(root);
    const workbook = path.join(root, 'Book.xlsm');
    fs.writeFileSync(workbook, '', 'utf8');
    return { root, workbook };
}

function testCase(): VbaTestCase {
    return {
        id: 'Tests.Test_Fail',
        moduleName: 'Tests',
        moduleType: 'standard',
        procedureName: 'Test_Fail',
        qualifiedName: 'Tests.Test_Fail',
        line: 4,
        column: 1,
        annotationLine: 3,
        metadata: {
            tags: ['ci'],
        },
    };
}

function reportFor(workbook: string): VbaTestRunReport {
    const test = testCase();
    return {
        filePath: workbook,
        workbookName: path.basename(workbook),
        startedAt: '2026-06-03T12:34:56.000Z',
        durationMs: 42,
        discovery: {
            filePath: workbook,
            tests: [test],
            unfilteredTestCount: 1,
            modulesScanned: 1,
            modulesIgnored: 0,
            contract: 'Standard-module no-argument Sub procedures with @xlide-test directives.',
        },
        results: [{
            test,
            status: 'failed',
            durationMs: 5,
            error: 'Expected <True> but was <False>.',
        }],
    };
}

describe('agent VBA test artifacts', () => {
    it('writes the same CI artifact surface for agent-driven test runs', async () => {
        const { workbook } = tempWorkbook();
        await writeWorkbookSettings(workbook, {
            tests: {
                artifactFolder: 'ci-tests',
                artifactRetention: 3,
            },
        });

        const artifacts = await writeAgentVbaTestArtifacts(reportFor(workbook), []);

        expect(artifacts.ok).toBe(true);
        if (!artifacts.ok) {
            throw new Error(artifacts.error);
        }
        expect(artifacts.runId).toBe('Book_2026-06-03_123456');
        expect(path.basename(artifacts.statusPath)).toBe('status_for_ci.json');
        expect(artifacts.settings).toMatchObject({
            artifactFolder: 'ci-tests',
            artifactFolderSource: 'workbook',
            artifactRetention: 3,
            artifactRetentionSource: 'workbook',
        });
        expect(fs.existsSync(artifacts.summaryPath)).toBe(true);
        expect(fs.existsSync(artifacts.hostTracePath)).toBe(true);
        expect(fs.existsSync(artifacts.outputLogPath)).toBe(true);
        expect(fs.existsSync(artifacts.statusPath)).toBe(true);
        expect(artifacts.ciStatus).toMatchObject({
            status: 'fail',
            reason: 'test-failures',
            counts: {
                total: 1,
                failed: 1,
            },
            failedTests: [{
                id: 'Tests.Test_Fail',
                module: 'Tests',
                procedure: 'Test_Fail',
                status: 'failed',
                message: 'Expected <True> but was <False>.',
            }],
        });
    });
});
