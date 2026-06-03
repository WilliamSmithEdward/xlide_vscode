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
                "' @xlide-test tags=smoke,fast owner=finance requirement=INV-104 timeout=2s expected-error=13",
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
            metadata: {
                tags: ['smoke', 'fast'],
                owner: 'finance',
                requirement: 'INV-104',
                timeoutMs: 2000,
                expectedError: '13',
            },
        })]);
    });

    it('discovers skip and expected-failure metadata from the annotation block', () => {
        const tests = discoverVbaTestsFromModule({
            name: 'MetadataTests',
            type: 'standard',
            source: [
                "' @xlide-test tags=known-bug",
                "' @xlide-test-xfail reason=\"Pending fix\"",
                'Sub KnownFailure()',
                'End Sub',
                '',
                "' @xlide-test-skip reason=\"Needs external workbook\"",
                'Sub ExternalScenario()',
                'End Sub',
            ].join('\n'),
        });

        expect(tests.map((test) => [test.qualifiedName, test.metadata])).toEqual([
            ['MetadataTests.KnownFailure', {
                tags: ['known-bug'],
                xfailReason: 'Pending fix',
            }],
            ['MetadataTests.ExternalScenario', {
                tags: [],
                skipReason: 'Needs external workbook',
            }],
        ]);
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
        expect(result.unfilteredTestCount).toBe(1);
        expect(result.tests.map((test) => test.qualifiedName)).toEqual(['Tests.TestWorkbookFlow']);
        expect(result.contract).toContain('@xlide-test');
    });

    it('filters discovered workbook tests by module, procedure, and tags', async () => {
        const bridge = bridgeForModules([
            {
                name: 'AlphaTests',
                type: 'standard',
                source: [
                    "' @xlide-test tags=smoke,fast",
                    'Sub FastSmoke()',
                    'End Sub',
                    '',
                    "' @xlide-test tags=slow",
                    'Sub SlowScenario()',
                    'End Sub',
                ].join('\n'),
            },
            {
                name: 'BetaTests',
                type: 'standard',
                source: [
                    "' @xlide-test tags=smoke",
                    'Sub BetaSmoke()',
                    'End Sub',
                ].join('\n'),
            },
        ]);

        const moduleResult = await discoverWorkbookVbaTests(bridge, 'C:/work/Book.xlsm', {
            moduleName: 'alphatests',
            includeTags: ['SMOKE'],
            excludeTags: ['slow'],
        });
        expect(moduleResult.modulesScanned).toBe(1);
        expect(moduleResult.unfilteredTestCount).toBe(2);
        expect(moduleResult.selection).toEqual({
            moduleName: 'alphatests',
            includeTags: ['SMOKE'],
            excludeTags: ['slow'],
        });
        expect(moduleResult.tests.map((test) => test.qualifiedName)).toEqual(['AlphaTests.FastSmoke']);

        const procedureResult = await discoverWorkbookVbaTests(bridge, 'C:/work/Book.xlsm', {
            moduleName: 'BetaTests',
            procedureName: 'betasmoke',
        });
        expect(procedureResult.unfilteredTestCount).toBe(1);
        expect(procedureResult.tests.map((test) => test.qualifiedName)).toEqual(['BetaTests.BetaSmoke']);
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
            metadata: { tags: [] },
        };
        const report = createVbaTestRunReport({
            filePath: 'C:/work/Book.xlsm',
            startedAt: new Date('2026-01-01T00:00:00Z'),
            durationMs: 12,
            discovery: {
                filePath: 'C:/work/Book.xlsm',
                tests: [test],
                unfilteredTestCount: 1,
                modulesScanned: 1,
                modulesIgnored: 0,
                contract: 'contract',
            },
            results: [
                { test, status: 'passed', durationMs: 3 },
                { test, status: 'failed', durationMs: 4, error: 'boom' },
                { test, status: 'skipped', durationMs: 0, error: 'not supported' },
                { test, status: 'xfail', durationMs: 5, error: 'known failure' },
                { test, status: 'xpass', durationMs: 6, error: 'unexpected pass' },
                { test, status: 'timeout', durationMs: 30000, error: 'hung' },
                { test, status: 'host-error', durationMs: 0, error: 'open failed' },
            ],
        });

        expect(report.workbookName).toBe('Book.xlsm');
        expect(summarizeVbaTestRun(report)).toEqual({
            total: 7,
            passed: 1,
            failed: 1,
            skipped: 1,
            xfail: 1,
            xpass: 1,
            timeout: 1,
            hostError: 1,
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
