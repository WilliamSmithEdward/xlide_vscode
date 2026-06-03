export const DEFAULT_VBA_TEST_HOST_ORACLE_CONTRACT = {
    excelInstance: 'single-owned-instance',
    workbookOpenMode: 'read-only',
    attachToUserExcelByDefault: false,
    closeWorkbookWithoutSaving: true,
    requiresPerTestTimeout: true,
    cleanupOwnedExcelOnHang: true,
    detectsExcelModals: true,
    dismissesSafeModalsWithoutSendKeys: true,
    blocksUnsafeModals: true,
} as const;

export type VbaTestHostOracleIssueCode =
    | 'empty-trace'
    | 'single-owned-excel-instance'
    | 'attached-excel-instance'
    | 'workbook-open-count'
    | 'workbook-open-instance'
    | 'read-only-workbook'
    | 'suppress-link-update'
    | 'suppress-alerts'
    | 'ignore-read-only-recommended'
    | 'macro-instance'
    | 'macro-order'
    | 'macro-timeout'
    | 'modal-result'
    | 'modal-cleanup'
    | 'close-without-saving'
    | 'normal-cleanup'
    | 'hang-cleanup'
    | 'no-macros-after-kill';

export interface VbaTestHostOracleIssue {
    code: VbaTestHostOracleIssueCode;
    message: string;
    eventIndex?: number;
}

export type VbaTestMacroOutcome = 'passed' | 'failed' | 'timeout' | 'hung' | 'modal-blocked' | 'runner-error';
export type VbaTestHostPhase =
    | 'excel-create'
    | 'workbook-open'
    | 'workbook-close'
    | 'excel-quit'
    | 'com-release';

export type VbaTestHostOracleEvent =
    | { kind: 'excel-created'; excelId: string; owned: boolean; pid?: number; visible?: boolean }
    | { kind: 'excel-attached'; excelId: string }
    | {
        kind: 'host-phase';
        excelId: string;
        phase: VbaTestHostPhase;
        outcome: 'passed' | 'failed';
        durationMs: number;
        message?: string;
    }
    | {
        kind: 'workbook-opened';
        excelId: string;
        filePath: string;
        readOnly: boolean;
        updateLinks?: number | boolean;
        displayAlerts?: boolean;
        ignoreReadOnlyRecommended?: boolean;
    }
    | { kind: 'macro-started'; excelId: string; qualifiedName: string; timeoutMs?: number }
    | {
        kind: 'modal-detected';
        excelId: string;
        qualifiedName: string;
        title?: string;
        className?: string;
        message?: string;
        texts?: string[];
        buttons?: string[];
        buttonIds?: number[];
        safeToDismiss?: boolean;
        classification?: string;
    }
    | {
        kind: 'modal-dismissed';
        excelId: string;
        qualifiedName: string;
        title?: string;
        message?: string;
        button?: string;
        buttonId?: number;
        dismissed: boolean;
    }
    | {
        kind: 'modal-blocked';
        excelId: string;
        qualifiedName: string;
        title?: string;
        message?: string;
        buttons?: string[];
        buttonIds?: number[];
        reason: string;
    }
    | {
        kind: 'macro-finished';
        excelId: string;
        qualifiedName: string;
        outcome: VbaTestMacroOutcome;
        durationMs?: number;
        message?: string;
    }
    | { kind: 'workbook-closed'; excelId: string; filePath?: string; saveChanges: boolean; durationMs?: number }
    | { kind: 'excel-quit'; excelId: string; durationMs?: number }
    | { kind: 'excel-killed'; excelId: string; reason: 'timeout' | 'hung' | 'modal-blocked' | 'runner-error' | 'cleanup-failed' };

export function validateVbaTestHostOracleTrace(
    events: readonly VbaTestHostOracleEvent[],
): VbaTestHostOracleIssue[] {
    const issues: VbaTestHostOracleIssue[] = [];
    if (events.length === 0) {
        return [{
            code: 'empty-trace',
            message: 'The test-host oracle trace must include the Excel lifecycle for a run.',
        }];
    }

    const created = indexed(events, 'excel-created');
    const attached = indexed(events, 'excel-attached');
    const opened = indexed(events, 'workbook-opened');
    const macroStarted = indexed(events, 'macro-started');
    const macroFinished = indexed(events, 'macro-finished');
    const modalBlocked = indexed(events, 'modal-blocked');
    const closed = indexed(events, 'workbook-closed');
    const quit = indexed(events, 'excel-quit');
    const killed = indexed(events, 'excel-killed');

    if (attached.length > 0) {
        for (const entry of attached) {
            issues.push({
                code: 'attached-excel-instance',
                message: 'The default VBA test host must not attach to a user Excel instance.',
                eventIndex: entry.index,
            });
        }
    }

    if (created.length !== 1 || !created[0]?.event.owned) {
        issues.push({
            code: 'single-owned-excel-instance',
            message: 'The default VBA test host must create exactly one XLIDE-owned Excel instance per run.',
            eventIndex: created[0]?.index,
        });
    }
    const excelId = created[0]?.event.excelId;

    if (opened.length !== 1) {
        issues.push({
            code: 'workbook-open-count',
            message: 'The default VBA test host must open exactly one workbook for the run.',
            eventIndex: opened[0]?.index,
        });
    }
    const openEntry = opened[0];
    if (openEntry && excelId && openEntry.event.excelId !== excelId) {
        issues.push({
            code: 'workbook-open-instance',
            message: 'The workbook must open inside the single XLIDE-owned Excel instance.',
            eventIndex: openEntry.index,
        });
    }
    if (openEntry) {
        if (!openEntry.event.readOnly) {
            issues.push({
                code: 'read-only-workbook',
                message: 'The default VBA test host must open the workbook read-only.',
                eventIndex: openEntry.index,
            });
        }
        if (openEntry.event.updateLinks !== 0 && openEntry.event.updateLinks !== false) {
            issues.push({
                code: 'suppress-link-update',
                message: 'The default VBA test host must disable link updates when opening the workbook.',
                eventIndex: openEntry.index,
            });
        }
        if (openEntry.event.displayAlerts !== false) {
            issues.push({
                code: 'suppress-alerts',
                message: 'The default VBA test host must suppress Excel alerts that can block automation.',
                eventIndex: openEntry.index,
            });
        }
        if (openEntry.event.ignoreReadOnlyRecommended !== true) {
            issues.push({
                code: 'ignore-read-only-recommended',
                message: 'The default VBA test host must bypass read-only recommendation prompts.',
                eventIndex: openEntry.index,
            });
        }
    }

    const openIndex = openEntry?.index ?? -1;
    const firstCloseOrKillIndex = firstIndexAfter(events, openIndex, ['workbook-closed', 'excel-killed']);
    for (const entry of macroStarted) {
        if (excelId && entry.event.excelId !== excelId) {
            issues.push({
                code: 'macro-instance',
                message: 'Every VBA test macro must run in the single XLIDE-owned Excel instance.',
                eventIndex: entry.index,
            });
        }
        if (entry.index <= openIndex || (firstCloseOrKillIndex >= 0 && entry.index > firstCloseOrKillIndex)) {
            issues.push({
                code: 'macro-order',
                message: 'VBA test macros must run after workbook open and before close or kill cleanup.',
                eventIndex: entry.index,
            });
        }
        const timeoutMs = entry.event.timeoutMs;
        if (timeoutMs === undefined || !Number.isFinite(timeoutMs) || timeoutMs <= 0) {
            issues.push({
                code: 'macro-timeout',
                message: 'Every VBA test macro must carry a positive timeout so hangs are bounded.',
                eventIndex: entry.index,
            });
        }
    }
    for (const entry of macroFinished) {
        if (excelId && entry.event.excelId !== excelId) {
            issues.push({
                code: 'macro-instance',
                message: 'Every VBA test macro result must come from the single XLIDE-owned Excel instance.',
                eventIndex: entry.index,
            });
        }
    }
    for (const entry of modalBlocked) {
        const resultAfterModal = macroFinished.find((finished) =>
            finished.index > entry.index &&
            finished.event.excelId === entry.event.excelId &&
            finished.event.qualifiedName === entry.event.qualifiedName &&
            finished.event.outcome === 'modal-blocked',
        );
        if (!resultAfterModal) {
            issues.push({
                code: 'modal-result',
                message: 'A blocked Excel modal must be reflected as a modal-blocked macro result.',
                eventIndex: entry.index,
            });
            continue;
        }
        const killAfterModal = killed.find((kill) =>
            kill.index > resultAfterModal.index && kill.event.excelId === entry.event.excelId,
        );
        if (!killAfterModal) {
            issues.push({
                code: 'modal-cleanup',
                message: 'A blocked Excel modal must clean up the XLIDE-owned Excel instance.',
                eventIndex: resultAfterModal.index,
            });
        }
    }

    const firstHang = macroFinished.find((entry) =>
        entry.event.outcome === 'timeout' || entry.event.outcome === 'hung' || entry.event.outcome === 'modal-blocked',
    );
    if (firstHang) {
        const killAfterHang = killed.find((entry) =>
            entry.event.excelId === firstHang.event.excelId && entry.index > firstHang.index,
        );
        if (!killAfterHang) {
            issues.push({
                code: 'hang-cleanup',
                message: 'A timeout or hang must clean up the XLIDE-owned Excel instance.',
                eventIndex: firstHang.index,
            });
        }
        const macroAfterKill = killAfterHang
            ? macroStarted.find((entry) => entry.index > killAfterHang.index)
            : undefined;
        if (macroAfterKill) {
            issues.push({
                code: 'no-macros-after-kill',
                message: 'No further VBA test macros may run after the owned Excel instance is killed.',
                eventIndex: macroAfterKill.index,
            });
        }
        return issues;
    }

    const closeEntry = closed[0];
    if (!closeEntry || closeEntry.event.excelId !== excelId || closeEntry.event.saveChanges) {
        issues.push({
            code: 'close-without-saving',
            message: 'Normal VBA test runs must close the workbook without saving changes.',
            eventIndex: closeEntry?.index,
        });
    }
    const quitAfterClose = closeEntry
        ? quit.find((entry) => entry.event.excelId === closeEntry.event.excelId && entry.index > closeEntry.index)
        : undefined;
    if (!quitAfterClose) {
        issues.push({
            code: 'normal-cleanup',
            message: 'Normal VBA test runs must quit the XLIDE-owned Excel instance after closing the workbook.',
            eventIndex: closeEntry?.index,
        });
    }

    return issues;
}

function indexed<K extends VbaTestHostOracleEvent['kind']>(
    events: readonly VbaTestHostOracleEvent[],
    kind: K,
): Array<{ index: number; event: Extract<VbaTestHostOracleEvent, { kind: K }> }> {
    const out: Array<{ index: number; event: Extract<VbaTestHostOracleEvent, { kind: K }> }> = [];
    events.forEach((event, index) => {
        if (event.kind === kind) {
            out.push({ index, event: event as Extract<VbaTestHostOracleEvent, { kind: K }> });
        }
    });
    return out;
}

function firstIndexAfter(
    events: readonly VbaTestHostOracleEvent[],
    index: number,
    kinds: readonly VbaTestHostOracleEvent['kind'][],
): number {
    return events.findIndex((event, eventIndex) =>
        eventIndex > index && kinds.includes(event.kind),
    );
}
