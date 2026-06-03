import { describe, expect, it, vi } from 'vitest';
import type { PythonBridge } from '../src/pythonBridge';
import { getVbaTestSupportStatus } from '../src/vbaTestSupportStatus';
import { XLIDE_ASSERT_MODULE_SOURCE } from '../src/vbaTestSupportModule';

function bridgeWithModules(
    modules: Array<{ name: string; type: string }>,
    sourceByModule: Record<string, string> = {},
): PythonBridge {
    return {
        call: vi.fn(async (method: string, args: { module?: string }) => {
            if (method === 'listModules') {
                return modules;
            }
            if (method === 'readModule' && args.module) {
                return { source: sourceByModule[args.module] ?? '' };
            }
            throw new Error(`Unexpected bridge call: ${method}`);
        }),
    } as unknown as PythonBridge;
}

describe('VBA test support status', () => {
    it('blocks test execution when XlideAssert is missing', async () => {
        const status = await getVbaTestSupportStatus(bridgeWithModules([]), 'Book.xlsm');

        expect(status).toEqual(expect.objectContaining({
            state: 'missing',
            canRun: false,
            canInstall: true,
        }));
    });

    it('allows test execution when the bundled support module is installed', async () => {
        const status = await getVbaTestSupportStatus(
            bridgeWithModules(
                [{ name: 'XlideAssert', type: 'standard' }],
                { XlideAssert: XLIDE_ASSERT_MODULE_SOURCE },
            ),
            'Book.xlsm',
        );

        expect(status).toEqual(expect.objectContaining({
            state: 'installed',
            canRun: true,
            canInstall: false,
        }));
    });

    it('blocks test execution when XlideAssert is outdated', async () => {
        const status = await getVbaTestSupportStatus(
            bridgeWithModules(
                [{ name: 'XlideAssert', type: 'standard' }],
                { XlideAssert: 'Option Explicit\n' },
            ),
            'Book.xlsm',
        );

        expect(status).toEqual(expect.objectContaining({
            state: 'outdated',
            canRun: false,
            canInstall: true,
        }));
    });
});
