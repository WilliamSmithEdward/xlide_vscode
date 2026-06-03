import { describe, expect, it } from 'vitest';
import type { PythonBridge } from '../src/pythonBridge';
import {
    createVbaTestRunReport,
    discoverVbaTestsFromModule,
    discoverWorkbookVbaTests,
    summarizeVbaTestRun,
    vbaTestFailureMessage,
} from '../src/vbaTestRunner';

describe('VBA test runner discovery', () => {
    it('discovers only explicitly annotated no-argument standard-module Subs', () => {
        const tests = discoverVbaTestsFromModule({
            name: 'TestModule',
            type: 'standard',
            source: [
                'Option Explicit',
                "' regular comment in the annotation block",
                "' @xlide-test",
                'Public Sub AddsNumbers()',
                'End Sub',
                '',
                "' @xlide-test",
                'Public Function NotRunnable() As Boolean',
                'End Function',
                '',
                "' @xlide-test",
                'Public Sub NeedsArg(value As Long)',
                'End Sub',
                '',
                'Public Sub TestNameOnly()',
                'End Sub',
            ].join('\n'),
        });

        expect(tests).toEqual([expect.objectContaining({
            id: 'TestModule.AddsNumbers',
            moduleName: 'TestModule',
            moduleType: 'standard',
            procedureName: 'AddsNumbers',
            qualifiedName: 'TestModule.AddsNumbers',
            line: 4,
            column: 1,
            annotationLine: 3,
        })]);
    });

    it('ignores annotations in modules Excel cannot run as standard macros', async () => {
        const bridge = bridgeForModules([
            {
                name: 'Tests',
                type: 'standard',
                source: "' @xlide-test\nSub TestWorkbookFlow()\nEnd Sub\n",
            },
            {
                name: 'Sheet1',
                type: 'document',
                source: "' @xlide-test\nSub TestSheetFlow()\nEnd Sub\n",
            },
            {
                name: 'Person',
                type: 'class',
                source: "' @xlide-test\nSub TestPersonFlow()\nEnd Sub\n",
            },
        ]);

        const result = await discoverWorkbookVbaTests(bridge, 'C:/work/Book.xlsm');

        expect(result.modulesScanned).toBe(1);
        expect(result.modulesIgnored).toBe(2);
        expect(result.tests.map((test) => test.qualifiedName)).toEqual(['Tests.TestWorkbookFlow']);
        expect(result.contract).toContain('@xlide-test');
    });
});

describe('VBA test runner reporting', () => {
    it('summarizes pass, fail, and skipped results', () => {
        const test = {
            id: 'Tests.TestA',
            moduleName: 'Tests',
            moduleType: 'standard',
            procedureName: 'TestA',
            qualifiedName: 'Tests.TestA',
            line: 2,
            column: 1,
            annotationLine: 1,
        };
        const report = createVbaTestRunReport({
            filePath: 'C:/work/Book.xlsm',
            startedAt: new Date('2026-01-01T00:00:00Z'),
            durationMs: 12,
            discovery: {
                filePath: 'C:/work/Book.xlsm',
                tests: [test],
                modulesScanned: 1,
                modulesIgnored: 0,
                contract: 'contract',
            },
            results: [
                { test, status: 'passed', durationMs: 3 },
                { test, status: 'failed', durationMs: 4, error: 'boom' },
                { test, status: 'skipped', durationMs: 0, error: 'not supported' },
            ],
        });

        expect(report.workbookName).toBe('Book.xlsm');
        expect(summarizeVbaTestRun(report)).toEqual({
            total: 3,
            passed: 1,
            failed: 1,
            skipped: 1,
        });
        expect(vbaTestFailureMessage(new Error('RUN_FAILED|Assertion failed'))).toBe('Assertion failed');
    });
});

function bridgeForModules(modules: Array<{ name: string; type: string; source: string }>): PythonBridge {
    return {
        async call<T>(method: string, params: { module?: string }): Promise<T> {
            if (method === 'listModules') {
                return modules.map(({ name, type }) => ({ name, type })) as T;
            }
            if (method === 'readModule' && params.module) {
                const module = modules.find((candidate) => candidate.name === params.module);
                if (!module) {
                    throw new Error(`Unknown module ${params.module}`);
                }
                return { source: module.source } as T;
            }
            throw new Error(`Unexpected bridge call ${method}`);
        },
    } as unknown as PythonBridge;
}
