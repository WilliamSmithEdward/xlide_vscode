import { describe, expect, it } from 'vitest';
import type { PythonBridge } from '../src/pythonBridge';
import {
    createVbaTestRunReport,
    describeVbaTestSelection,
    VBA_TEST_DIRECTIVE_DIAGNOSTIC_CODE,
    discoverVbaTestsFromModule,
    discoverWorkbookVbaTests,
    summarizeVbaTestTags,
    summarizeVbaTestRun,
    validateVbaTestDirectivesFromModule,
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

    it('discovers expected-error metadata for any caught VBA error', () => {
        const tests = discoverVbaTestsFromModule({
            name: 'ExpectedErrorTests',
            type: 'standard',
            source: [
                "' @xlide-test expected-error",
                'Sub AnyErrorBare()',
                'End Sub',
                '',
                "' @xlide-test expected-error=any",
                'Sub AnyErrorValue()',
                'End Sub',
            ].join('\n'),
        });

        expect(tests.map((test) => [test.procedureName, test.metadata.expectedError])).toEqual([
            ['AnyErrorBare', 'any'],
            ['AnyErrorValue', 'any'],
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

    it('discovers workbook tests through a single batch bridge read', async () => {
        const calls: string[] = [];
        const bridge = {
            async call<T>(method: string, params: { module?: string }): Promise<T> {
                calls.push(params.module ? `${method}:${params.module}` : method);
                if (method === 'readModules') {
                    return [
                        { name: 'Tests', type: 'standard', source: "' @xlide-test\nSub Runs()\nEnd Sub\n" },
                        { name: 'Sheet1', type: 'document', source: '' },
                    ] as T;
                }
                throw new Error(`Unexpected bridge call ${method}`);
            },
        } as unknown as PythonBridge;

        const result = await discoverWorkbookVbaTests(bridge, 'C:/work/Book.xlsm');

        expect(result.tests.map((test) => test.qualifiedName)).toEqual(['Tests.Runs']);
        expect(calls).toEqual(['readModules']);
    });

    it('falls back to per-module bridge reads when the backend lacks readModules', async () => {
        const calls: string[] = [];
        const bridge = {
            async call<T>(method: string, params: { module?: string }): Promise<T> {
                calls.push(params.module ? `${method}:${params.module}` : method);
                if (method === 'readModules') {
                    throw new Error('Method not found: readModules');
                }
                if (method === 'listModules') {
                    return [
                        { name: 'Tests', type: 'standard' },
                        { name: 'Sheet1', type: 'document' },
                    ] as T;
                }
                if (method === 'readModule' && params.module === 'Tests') {
                    return { source: "' @xlide-test\nSub Runs()\nEnd Sub\n" } as T;
                }
                throw new Error(`Unexpected bridge call ${method}`);
            },
        } as unknown as PythonBridge;

        const result = await discoverWorkbookVbaTests(bridge, 'C:/work/Book.xlsm');

        expect(result.tests.map((test) => test.qualifiedName)).toEqual(['Tests.Runs']);
        expect(calls).toEqual(['readModules', 'listModules', 'readModule:Tests']);
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

        const idResult = await discoverWorkbookVbaTests(bridge, 'C:/work/Book.xlsm', {
            testIds: ['betatests.betasmoke', 'AlphaTests.FastSmoke'],
        });
        expect(idResult.selection).toEqual({
            testIds: ['betatests.betasmoke', 'AlphaTests.FastSmoke'],
        });
        expect(idResult.tests.map((test) => test.qualifiedName)).toEqual([
            'AlphaTests.FastSmoke',
            'BetaTests.BetaSmoke',
        ]);
        expect(describeVbaTestSelection(idResult.selection)).toBe('2 selected tests');
    });

    it('summarizes discovered tags by name for workbook filter UI', () => {
        const tests = discoverVbaTestsFromModule({
            name: 'TagTests',
            type: 'standard',
            source: [
                "' @xlide-test tags=Smoke,fast,smoke",
                'Sub SmokeFast()',
                'End Sub',
                '',
                "' @xlide-test tags=slow",
                'Sub SlowScenario()',
                'End Sub',
            ].join('\n'),
        });

        expect(summarizeVbaTestTags(tests)).toEqual([
            { name: 'fast', testCount: 1 },
            { name: 'slow', testCount: 1 },
            { name: 'smoke', testCount: 1 },
        ]);
    });

    it('validates malformed test directive syntax and metadata', () => {
        const issues = validateVbaTestDirectivesFromModule({
            name: 'Tests',
            type: 'standard',
            source: [
                "' @xlide-tset tags=smoke",
                'Sub Typo()',
                'End Sub',
                '',
                "' @xlide-test timeout=fast owner finance",
                'Sub BadMetadata()',
                'End Sub',
                '',
                "' @xlide-test-skip",
                'Sub MissingReason()',
                'End Sub',
                '',
                "' @xlide-test tags=,",
                'Sub EmptyTags()',
                'End Sub',
                '',
                "' @xlide-test expected-error=abc",
                'Sub BadExpectedError()',
                'End Sub',
            ].join('\n'),
        });

        expect(issues.map((issue) => issue.code)).toEqual(Array(6).fill(VBA_TEST_DIRECTIVE_DIAGNOSTIC_CODE));
        expect(issues.map((issue) => `${issue.line}:${issue.message}`)).toEqual([
            '1:Unknown XLIDE test directive. Supported directives are @xlide-test, @xlide-test-skip, and @xlide-test-xfail.',
            '5:Malformed XLIDE test metadata. Use key=value pairs and quote values that contain spaces.',
            '5:XLIDE test timeout must be a positive integer with optional ms or s suffix.',
            '9:XLIDE skip test directives should include reason="...".',
            '13:XLIDE test metadata key tags must list at least one tag.',
            '17:XLIDE expected-error metadata must be a positive VBA error number or any.',
        ]);
    });

    it('validates directives that cannot discover runnable tests', () => {
        const standardIssues = validateVbaTestDirectivesFromModule({
            name: 'Tests',
            type: 'standard',
            source: [
                "' @xlide-test",
                'Function NotRunnable() As Boolean',
                'End Function',
                '',
                "' @xlide-test",
                'Sub NeedsArg(value As Long)',
                'End Sub',
                '',
                "' @xlide-test",
                '',
                'Sub Detached()',
                'End Sub',
            ].join('\n'),
        });
        const classIssues = validateVbaTestDirectivesFromModule({
            name: 'Person',
            type: 'class',
            source: "' @xlide-test\nSub ClassScenario()\nEnd Sub",
        });

        expect(standardIssues.map((issue) => `${issue.line}:${issue.message}`)).toEqual([
            '1:XLIDE test directives must target a Sub procedure; Functions and Properties are not runnable tests.',
            '5:XLIDE test Sub procedures must not declare parameters.',
            '9:XLIDE test directives must be in the comment block immediately above a zero-argument Sub procedure.',
        ]);
        expect(classIssues.map((issue) => issue.message)).toEqual([
            "XLIDE test directives only run from standard modules; 'Person' is a class module.",
        ]);
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
        expect(vbaTestFailureMessage(new Error([
            'RUN_FAILED|Exception calling "Run" with "1" argument(s):',
            '"Exception from HRESULT: 0x800A9C68"',
            'HRESULT: 0x80131501',
            'Inner: Exception from HRESULT: 0x800A9C68',
            'Inner HRESULT: 0x800A9C68',
            'At C:\\Users\\William\\AppData\\Local\\Temp\\xlide-vba-test-host-DDT6Nq\\run-vba-tests.ps1:677 char:4587',
            '+ ... tPrefix, $excelId, $macroName) }; $excel.Run($macroRef);',
            '+                       ~~~~~~~~~~~~~~~~~~~~',
        ].join('\n')))).toBe(
            'Excel could not run the test macro. Check for VBA compile errors, macro security prompts, or a missing test procedure. HRESULT: 0x800A9C68.',
        );
        expect(vbaTestFailureMessage('A user message with | punctuation')).toBe('A user message with | punctuation');
        expect(vbaTestFailureMessage([
            'RUN_FAILED|HRESULT: 0x80131501',
            'Exception calling "Run" with "2" argument(s): "The remote procedure call failed. (Exception from HRESULT: 0x800706BE)"',
            'Inner: The remote procedure call failed. (Exception from HRESULT: 0x800706BE)',
            'Inner HRESULT: 0x800706BE',
        ].join('\n'))).toBe(
            'Excel automation became unavailable while running the test. Excel may have closed, crashed, or been blocked by a modal dialog. HRESULT: 0x800706BE.',
        );
        expect(vbaTestFailureMessage([
            'RUN_FAILED|HRESULT: 0x80131501',
            'Exception calling "Run" with "2" argument(s): "The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)"',
            'Inner HRESULT: 0x800706BA',
        ].join('\n'))).toBe(
            'Excel automation became unavailable while running the test. Excel may have closed, crashed, or been blocked by a modal dialog. HRESULT: 0x800706BA.',
        );
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
