import { describe, expect, it } from 'vitest';
import {
    validateVbaTestHostOracleTrace,
    type VbaTestHostOracleEvent,
} from '../src/vbaTestHostOracle';

describe('VBA test host oracle', () => {
    it('accepts one owned Excel instance, one read-only workbook, and normal cleanup', () => {
        const events: VbaTestHostOracleEvent[] = [
            { kind: 'host-created', hostId: 'xlide-1', owned: true },
            {
                kind: 'file-opened',
                hostId: 'xlide-1',
                filePath: 'C:/work/Book.xlsm',
                readOnly: true,
                updateLinks: 0,
                displayAlerts: false,
                ignoreReadOnlyRecommended: true,
            },
            { kind: 'macro-started', hostId: 'xlide-1', qualifiedName: 'Tests.Pass', timeoutMs: 5000 },
            { kind: 'macro-finished', hostId: 'xlide-1', qualifiedName: 'Tests.Pass', outcome: 'passed' },
            { kind: 'macro-started', hostId: 'xlide-1', qualifiedName: 'Tests.Fail', timeoutMs: 5000 },
            { kind: 'macro-finished', hostId: 'xlide-1', qualifiedName: 'Tests.Fail', outcome: 'failed' },
            { kind: 'file-closed', hostId: 'xlide-1', filePath: 'C:/work/Book.xlsm', saveChanges: false },
            { kind: 'host-quit', hostId: 'xlide-1' },
        ];

        expect(validateVbaTestHostOracleTrace(events)).toEqual([]);
    });

    it('rejects attaching to user Excel or creating multiple Excel instances', () => {
        const events: VbaTestHostOracleEvent[] = [
            { kind: 'host-attached', hostId: 'user-excel' },
            { kind: 'host-created', hostId: 'xlide-1', owned: true },
            { kind: 'host-created', hostId: 'xlide-2', owned: true },
            {
                kind: 'file-opened',
                hostId: 'xlide-1',
                filePath: 'C:/work/Book.xlsm',
                readOnly: true,
                updateLinks: 0,
                displayAlerts: false,
                ignoreReadOnlyRecommended: true,
            },
            { kind: 'file-closed', hostId: 'xlide-1', saveChanges: false },
            { kind: 'host-quit', hostId: 'xlide-1' },
        ];

        expect(issueCodes(events)).toEqual(expect.arrayContaining([
            'attached-host-instance',
            'single-owned-host-instance',
        ]));
    });

    it('rejects workbook opens that can mutate files or block automation', () => {
        const events: VbaTestHostOracleEvent[] = [
            { kind: 'host-created', hostId: 'xlide-1', owned: true },
            {
                kind: 'file-opened',
                hostId: 'xlide-1',
                filePath: 'C:/work/Book.xlsm',
                readOnly: false,
                updateLinks: true,
                displayAlerts: true,
                ignoreReadOnlyRecommended: false,
            },
            { kind: 'file-closed', hostId: 'xlide-1', saveChanges: true },
        ];

        expect(issueCodes(events)).toEqual(expect.arrayContaining([
            'read-only-file',
            'suppress-link-update',
            'suppress-alerts',
            'ignore-read-only-recommended',
            'close-without-saving',
            'normal-cleanup',
        ]));
    });

    it('requires timeouts and owned Excel cleanup after hangs', () => {
        const missingCleanup: VbaTestHostOracleEvent[] = [
            { kind: 'host-created', hostId: 'xlide-1', owned: true },
            {
                kind: 'file-opened',
                hostId: 'xlide-1',
                filePath: 'C:/work/Book.xlsm',
                readOnly: true,
                updateLinks: 0,
                displayAlerts: false,
                ignoreReadOnlyRecommended: true,
            },
            { kind: 'macro-started', hostId: 'xlide-1', qualifiedName: 'Tests.Hangs' },
            { kind: 'macro-finished', hostId: 'xlide-1', qualifiedName: 'Tests.Hangs', outcome: 'timeout' },
        ];

        expect(issueCodes(missingCleanup)).toEqual(expect.arrayContaining([
            'macro-timeout',
            'hang-cleanup',
        ]));

        const cleanedUp: VbaTestHostOracleEvent[] = [
            ...missingCleanup.slice(0, 2),
            { kind: 'macro-started', hostId: 'xlide-1', qualifiedName: 'Tests.Hangs', timeoutMs: 5000 },
            { kind: 'macro-finished', hostId: 'xlide-1', qualifiedName: 'Tests.Hangs', outcome: 'hung' },
            { kind: 'host-killed', hostId: 'xlide-1', reason: 'hung' },
        ];
        expect(validateVbaTestHostOracleTrace(cleanedUp)).toEqual([]);
    });

    it('accepts informational modal detection and dismissal during a normal run', () => {
        const events: VbaTestHostOracleEvent[] = [
            { kind: 'host-created', hostId: 'xlide-1', owned: true },
            {
                kind: 'file-opened',
                hostId: 'xlide-1',
                filePath: 'C:/work/Book.xlsm',
                readOnly: true,
                updateLinks: 0,
                displayAlerts: false,
                ignoreReadOnlyRecommended: true,
            },
            { kind: 'macro-started', hostId: 'xlide-1', qualifiedName: 'Tests.MsgBox', timeoutMs: 5000 },
            {
                kind: 'modal-detected',
                hostId: 'xlide-1',
                qualifiedName: 'Tests.MsgBox',
                title: 'XLIDE Modal Smoke',
                message: 'XLIDE modal smoke',
                buttons: ['OK'],
                buttonIds: [1],
                safeToDismiss: true,
                classification: 'host-modal',
            },
            {
                kind: 'modal-dismissed',
                hostId: 'xlide-1',
                qualifiedName: 'Tests.MsgBox',
                title: 'XLIDE Modal Smoke',
                button: 'OK',
                buttonId: 1,
                dismissed: true,
            },
            { kind: 'macro-finished', hostId: 'xlide-1', qualifiedName: 'Tests.MsgBox', outcome: 'passed' },
            { kind: 'file-closed', hostId: 'xlide-1', filePath: 'C:/work/Book.xlsm', saveChanges: false },
            { kind: 'host-quit', hostId: 'xlide-1' },
        ];

        expect(validateVbaTestHostOracleTrace(events)).toEqual([]);
    });

    it('accepts multiple safe modal dialogs in a single macro', () => {
        const events: VbaTestHostOracleEvent[] = [
            { kind: 'host-created', hostId: 'xlide-1', owned: true },
            {
                kind: 'file-opened',
                hostId: 'xlide-1',
                filePath: 'C:/work/Book.xlsm',
                readOnly: true,
                updateLinks: 0,
                displayAlerts: false,
                ignoreReadOnlyRecommended: true,
            },
            { kind: 'macro-started', hostId: 'xlide-1', qualifiedName: 'Tests.ChainedMsgBox', timeoutMs: 5000 },
            {
                kind: 'modal-detected',
                hostId: 'xlide-1',
                qualifiedName: 'Tests.ChainedMsgBox',
                title: 'XLIDE Chain',
                message: 'one',
                buttons: ['OK'],
                safeToDismiss: true,
            },
            {
                kind: 'modal-dismissed',
                hostId: 'xlide-1',
                qualifiedName: 'Tests.ChainedMsgBox',
                title: 'XLIDE Chain',
                button: 'OK',
                dismissed: true,
            },
            {
                kind: 'modal-detected',
                hostId: 'xlide-1',
                qualifiedName: 'Tests.ChainedMsgBox',
                title: 'XLIDE Chain',
                message: 'two',
                buttons: ['OK'],
                safeToDismiss: true,
            },
            {
                kind: 'modal-dismissed',
                hostId: 'xlide-1',
                qualifiedName: 'Tests.ChainedMsgBox',
                title: 'XLIDE Chain',
                button: 'OK',
                dismissed: true,
            },
            { kind: 'macro-finished', hostId: 'xlide-1', qualifiedName: 'Tests.ChainedMsgBox', outcome: 'passed' },
            { kind: 'file-closed', hostId: 'xlide-1', filePath: 'C:/work/Book.xlsm', saveChanges: false },
            { kind: 'host-quit', hostId: 'xlide-1' },
        ];

        expect(validateVbaTestHostOracleTrace(events)).toEqual([]);
    });

    it('requires blocked modal results to kill the owned Excel instance', () => {
        const missingResultAndCleanup: VbaTestHostOracleEvent[] = [
            { kind: 'host-created', hostId: 'xlide-1', owned: true },
            {
                kind: 'file-opened',
                hostId: 'xlide-1',
                filePath: 'C:/work/Book.xlsm',
                readOnly: true,
                updateLinks: 0,
                displayAlerts: false,
                ignoreReadOnlyRecommended: true,
            },
            { kind: 'macro-started', hostId: 'xlide-1', qualifiedName: 'Tests.DecisionDialog', timeoutMs: 5000 },
            {
                kind: 'modal-blocked',
                hostId: 'xlide-1',
                qualifiedName: 'Tests.DecisionDialog',
                title: 'Microsoft Excel',
                message: 'Save changes?',
                buttons: ['Yes', 'No', 'Cancel'],
                buttonIds: [6, 7, 2],
                reason: 'decision-or-unknown-dialog',
            },
        ];

        expect(issueCodes(missingResultAndCleanup)).toEqual(expect.arrayContaining([
            'modal-result',
            'close-without-saving',
            'normal-cleanup',
        ]));

        const cleanedUp: VbaTestHostOracleEvent[] = [
            ...missingResultAndCleanup,
            {
                kind: 'macro-finished',
                hostId: 'xlide-1',
                qualifiedName: 'Tests.DecisionDialog',
                outcome: 'modal-blocked',
                durationMs: 5000,
                message: 'Blocked by Excel modal dialog.',
            },
            { kind: 'host-killed', hostId: 'xlide-1', reason: 'modal-blocked' },
        ];
        expect(validateVbaTestHostOracleTrace(cleanedUp)).toEqual([]);
    });
});

function issueCodes(events: readonly VbaTestHostOracleEvent[]): string[] {
    return validateVbaTestHostOracleTrace(events).map((issue) => issue.code);
}
