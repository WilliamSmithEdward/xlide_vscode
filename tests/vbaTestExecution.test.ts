import { describe, expect, it } from 'vitest';
import {
    vbaTestRunItemFromHostResult,
    type OwnedReadOnlyExcelHostTestResult,
} from '../src/vbaTestExecution';
import type { VbaTestCase } from '../src/vbaTestRunner';

describe('VBA test execution result classification', () => {
    it('passes expected-error tests when the host reports the expected VBA error number', () => {
        const result = vbaTestRunItemFromHostResult(
            testCase({ expectedError: '13' }),
            hostResult({
                outcome: 'failed',
                errorNumber: 13,
                message: 'RUN_FAILED|VBA error 13 from VBAProject: Type mismatch',
            }),
        );

        expect(result.status).toBe('passed');
        expect(result.error).toBeUndefined();
    });

    it('fails expected-error tests when no error is raised', () => {
        const result = vbaTestRunItemFromHostResult(
            testCase({ expectedError: '13' }),
            hostResult({ outcome: 'passed' }),
        );

        expect(result.status).toBe('failed');
        expect(result.error).toBe('Expected VBA error 13, but no error was raised.');
    });

    it('passes expected-error any tests when the host reports any caught VBA error', () => {
        const result = vbaTestRunItemFromHostResult(
            testCase({ expectedError: 'any' }),
            hostResult({
                outcome: 'failed',
                errorNumber: 9,
                message: 'RUN_FAILED|VBA error 9 from VBAProject: Subscript out of range',
            }),
        );

        expect(result.status).toBe('passed');
        expect(result.error).toBeUndefined();
    });

    it('fails expected-error any tests when no error is raised', () => {
        const result = vbaTestRunItemFromHostResult(
            testCase({ expectedError: 'any' }),
            hostResult({ outcome: 'passed' }),
        );

        expect(result.status).toBe('failed');
        expect(result.error).toBe('Expected a VBA error, but no error was raised.');
    });

    it('fails expected-error tests when the host reports a different VBA error number', () => {
        const result = vbaTestRunItemFromHostResult(
            testCase({ expectedError: '13' }),
            hostResult({
                outcome: 'failed',
                errorNumber: 9,
                message: 'RUN_FAILED|VBA error 9 from VBAProject: Subscript out of range',
            }),
        );

        expect(result.status).toBe('failed');
        expect(result.error).toContain('Expected VBA error 13, but got VBA error 9');
    });

    it('treats expected-error success as unexpected pass when the test is also xfail', () => {
        const result = vbaTestRunItemFromHostResult(
            testCase({ expectedError: '13', xfailReason: 'Bug should still raise a different error' }),
            hostResult({ outcome: 'failed', errorNumber: 13 }),
        );

        expect(result.status).toBe('xpass');
        expect(result.error).toBe('Expected failure did not occur: Bug should still raise a different error');
    });

    it('preserves deterministic output lines from the VBA assertion API', () => {
        const passed = vbaTestRunItemFromHostResult(
            testCase(),
            hostResult({ outcome: 'passed', output: ['created invoice', 'checked total'] }),
        );
        const failed = vbaTestRunItemFromHostResult(
            testCase(),
            hostResult({
                outcome: 'failed',
                message: 'RUN_FAILED|Expected <1> but was <2>.',
                output: ['before assertion'],
            }),
        );

        expect(passed).toMatchObject({
            status: 'passed',
            output: ['created invoice', 'checked total'],
        });
        expect(failed).toMatchObject({
            status: 'failed',
            output: ['before assertion'],
        });
    });
});

function testCase(metadata: Partial<VbaTestCase['metadata']> = {}): VbaTestCase {
    return {
        id: 'Tests.ExpectedError',
        moduleName: 'Tests',
        moduleType: 'standard',
        procedureName: 'ExpectedError',
        qualifiedName: 'Tests.ExpectedError',
        line: 4,
        column: 1,
        annotationLine: 3,
        metadata: {
            tags: [],
            ...metadata,
        },
    };
}

function hostResult(input: Partial<OwnedReadOnlyExcelHostTestResult>): OwnedReadOnlyExcelHostTestResult {
    return {
        outcome: 'passed',
        durationMs: 12,
        ...input,
    };
}
