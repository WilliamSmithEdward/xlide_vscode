import { describe, expect, it } from 'vitest';
import {
    DEFAULT_VBA_TEST_HOST_ORACLE_CONTRACT,
    validateVbaTestHostOracleTrace,
    type VbaTestHostOracleEvent,
} from '../src/vbaTestHostOracle';

describe('VBA test host oracle', () => {
    it('accepts one owned Excel instance, one read-only workbook, and normal cleanup', () => {
        const events: VbaTestHostOracleEvent[] = [
            { kind: 'excel-created', excelId: 'xlide-1', owned: true },
            {
                kind: 'workbook-opened',
                excelId: 'xlide-1',
                filePath: 'C:/work/Book.xlsm',
                readOnly: true,
                updateLinks: 0,
                displayAlerts: false,
                ignoreReadOnlyRecommended: true,
            },
            { kind: 'macro-started', excelId: 'xlide-1', qualifiedName: 'Tests.Pass', timeoutMs: 5000 },
            { kind: 'macro-finished', excelId: 'xlide-1', qualifiedName: 'Tests.Pass', outcome: 'passed' },
            { kind: 'macro-started', excelId: 'xlide-1', qualifiedName: 'Tests.Fail', timeoutMs: 5000 },
            { kind: 'macro-finished', excelId: 'xlide-1', qualifiedName: 'Tests.Fail', outcome: 'failed' },
            { kind: 'workbook-closed', excelId: 'xlide-1', filePath: 'C:/work/Book.xlsm', saveChanges: false },
            { kind: 'excel-quit', excelId: 'xlide-1' },
        ];

        expect(DEFAULT_VBA_TEST_HOST_ORACLE_CONTRACT).toMatchObject({
            excelInstance: 'single-owned-instance',
            workbookOpenMode: 'read-only',
            attachToUserExcelByDefault: false,
        });
        expect(validateVbaTestHostOracleTrace(events)).toEqual([]);
    });

    it('rejects attaching to user Excel or creating multiple Excel instances', () => {
        const events: VbaTestHostOracleEvent[] = [
            { kind: 'excel-attached', excelId: 'user-excel' },
            { kind: 'excel-created', excelId: 'xlide-1', owned: true },
            { kind: 'excel-created', excelId: 'xlide-2', owned: true },
            {
                kind: 'workbook-opened',
                excelId: 'xlide-1',
                filePath: 'C:/work/Book.xlsm',
                readOnly: true,
                updateLinks: 0,
                displayAlerts: false,
                ignoreReadOnlyRecommended: true,
            },
            { kind: 'workbook-closed', excelId: 'xlide-1', saveChanges: false },
            { kind: 'excel-quit', excelId: 'xlide-1' },
        ];

        expect(issueCodes(events)).toEqual(expect.arrayContaining([
            'attached-excel-instance',
            'single-owned-excel-instance',
        ]));
    });

    it('rejects workbook opens that can mutate files or block automation', () => {
        const events: VbaTestHostOracleEvent[] = [
            { kind: 'excel-created', excelId: 'xlide-1', owned: true },
            {
                kind: 'workbook-opened',
                excelId: 'xlide-1',
                filePath: 'C:/work/Book.xlsm',
                readOnly: false,
                updateLinks: true,
                displayAlerts: true,
                ignoreReadOnlyRecommended: false,
            },
            { kind: 'workbook-closed', excelId: 'xlide-1', saveChanges: true },
        ];

        expect(issueCodes(events)).toEqual(expect.arrayContaining([
            'read-only-workbook',
            'suppress-link-update',
            'suppress-alerts',
            'ignore-read-only-recommended',
            'close-without-saving',
            'normal-cleanup',
        ]));
    });

    it('requires timeouts and owned Excel cleanup after hangs', () => {
        const missingCleanup: VbaTestHostOracleEvent[] = [
            { kind: 'excel-created', excelId: 'xlide-1', owned: true },
            {
                kind: 'workbook-opened',
                excelId: 'xlide-1',
                filePath: 'C:/work/Book.xlsm',
                readOnly: true,
                updateLinks: 0,
                displayAlerts: false,
                ignoreReadOnlyRecommended: true,
            },
            { kind: 'macro-started', excelId: 'xlide-1', qualifiedName: 'Tests.Hangs' },
            { kind: 'macro-finished', excelId: 'xlide-1', qualifiedName: 'Tests.Hangs', outcome: 'timeout' },
        ];

        expect(issueCodes(missingCleanup)).toEqual(expect.arrayContaining([
            'macro-timeout',
            'hang-cleanup',
        ]));

        const cleanedUp: VbaTestHostOracleEvent[] = [
            ...missingCleanup.slice(0, 2),
            { kind: 'macro-started', excelId: 'xlide-1', qualifiedName: 'Tests.Hangs', timeoutMs: 5000 },
            { kind: 'macro-finished', excelId: 'xlide-1', qualifiedName: 'Tests.Hangs', outcome: 'hung' },
            { kind: 'excel-killed', excelId: 'xlide-1', reason: 'hung' },
        ];
        expect(validateVbaTestHostOracleTrace(cleanedUp)).toEqual([]);
    });
});

function issueCodes(events: readonly VbaTestHostOracleEvent[]): string[] {
    return validateVbaTestHostOracleTrace(events).map((issue) => issue.code);
}
