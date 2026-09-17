export type VbaTestHostOracleIssueCode =
    | 'empty-trace'
    | 'single-owned-host-instance'
    | 'attached-host-instance'
    | 'file-open-count'
    | 'file-open-instance'
    | 'read-only-file'
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
    | 'host-create'
    | 'file-open'
    | 'file-close'
    | 'host-quit'
    | 'com-release';

export type VbaTestHostOracleEvent =
    | { kind: 'host-created'; hostId: string; owned: boolean; pid?: number; visible?: boolean }
    | { kind: 'host-attached'; hostId: string }
    | {
        kind: 'host-phase';
        hostId: string;
        phase: VbaTestHostPhase;
        outcome: 'passed' | 'failed';
        durationMs: number;
        message?: string;
    }
    | {
        kind: 'file-opened';
        hostId: string;
        filePath: string;
        readOnly: boolean;
        updateLinks?: number | boolean;
        displayAlerts?: boolean;
        ignoreReadOnlyRecommended?: boolean;
    }
    | { kind: 'macro-started'; hostId: string; qualifiedName: string; timeoutMs?: number }
    | {
        kind: 'modal-detected';
        hostId: string;
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
        hostId: string;
        qualifiedName: string;
        title?: string;
        message?: string;
        button?: string;
        buttonId?: number;
        dismissed: boolean;
    }
    | {
        kind: 'modal-blocked';
        hostId: string;
        qualifiedName: string;
        title?: string;
        message?: string;
        buttons?: string[];
        buttonIds?: number[];
        reason: string;
    }
    | {
        kind: 'macro-finished';
        hostId: string;
        qualifiedName: string;
        outcome: VbaTestMacroOutcome;
        durationMs?: number;
        message?: string;
        errorNumber?: number;
        errorSource?: string;
        output?: string[];
    }
    | { kind: 'file-closed'; hostId: string; filePath?: string; saveChanges: boolean; durationMs?: number }
    | { kind: 'host-quit'; hostId: string; durationMs?: number }
    | { kind: 'host-killed'; hostId: string; reason: 'timeout' | 'hung' | 'modal-blocked' | 'runner-error' | 'cleanup-failed' };

/** Prefix marking machine-readable oracle events on the PowerShell host's stdout. */
export const XLIDE_TEST_HOST_EVENT_PREFIX = 'XLIDE_TEST_HOST_EVENT|';

/** Parses one host stdout line; returns undefined for ordinary (non-event) output. */
export function parseVbaTestHostEventLine(line: string): VbaTestHostOracleEvent | undefined {
    if (!line.startsWith(XLIDE_TEST_HOST_EVENT_PREFIX)) {
        return undefined;
    }
    const json = line.slice(XLIDE_TEST_HOST_EVENT_PREFIX.length);
    const parsed: unknown = JSON.parse(json);
    // Validate the minimum shape: a structurally-wrong-but-valid-JSON object
    // must not be ingested into the session state machine / oracle as an event.
    if (typeof parsed !== 'object' || parsed === null
        || typeof (parsed as { kind?: unknown }).kind !== 'string') {
        return undefined;
    }
    return parsed as VbaTestHostOracleEvent;
}

export function validateVbaTestHostOracleTrace(
    events: readonly VbaTestHostOracleEvent[],
): VbaTestHostOracleIssue[] {
    const issues: VbaTestHostOracleIssue[] = [];
    if (events.length === 0) {
        return [{
            code: 'empty-trace',
            message: 'The test-host oracle trace must include the host application\'s lifecycle for a run.',
        }];
    }

    const created = indexed(events, 'host-created');
    const attached = indexed(events, 'host-attached');
    const opened = indexed(events, 'file-opened');
    const macroStarted = indexed(events, 'macro-started');
    const macroFinished = indexed(events, 'macro-finished');
    const modalBlocked = indexed(events, 'modal-blocked');
    const closed = indexed(events, 'file-closed');
    const quit = indexed(events, 'host-quit');
    const killed = indexed(events, 'host-killed');

    if (attached.length > 0) {
        for (const entry of attached) {
            issues.push({
                code: 'attached-host-instance',
                message: 'The default VBA test host must not attach to an application instance the user is running.',
                eventIndex: entry.index,
            });
        }
    }

    if (created.length !== 1 || !created[0]?.event.owned) {
        issues.push({
            code: 'single-owned-host-instance',
            message: 'The default VBA test host must create exactly one XLIDE-owned application instance per run.',
            eventIndex: created[0]?.index,
        });
    }
    const hostId = created[0]?.event.hostId;

    if (opened.length !== 1) {
        issues.push({
            code: 'file-open-count',
            message: 'The default VBA test host must open exactly one file for the run.',
            eventIndex: opened[0]?.index,
        });
    }
    const openEntry = opened[0];
    if (openEntry && hostId && openEntry.event.hostId !== hostId) {
        issues.push({
            code: 'file-open-instance',
            message: 'The file must open inside the single XLIDE-owned application instance.',
            eventIndex: openEntry.index,
        });
    }
    if (openEntry) {
        if (!openEntry.event.readOnly) {
            issues.push({
                code: 'read-only-file',
                message: 'The default VBA test host must open the file read-only.',
                eventIndex: openEntry.index,
            });
        }
        if (openEntry.event.updateLinks !== 0 && openEntry.event.updateLinks !== false) {
            issues.push({
                code: 'suppress-link-update',
                message: 'The default VBA test host must disable link updates when opening the file.',
                eventIndex: openEntry.index,
            });
        }
        if (openEntry.event.displayAlerts !== false) {
            issues.push({
                code: 'suppress-alerts',
                message: 'The default VBA test host must suppress application alerts that can block automation.',
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
    const firstCloseOrKillIndex = firstIndexAfter(events, openIndex, ['file-closed', 'host-killed']);
    for (const entry of macroStarted) {
        if (hostId && entry.event.hostId !== hostId) {
            issues.push({
                code: 'macro-instance',
                message: 'Every VBA test macro must run in the single XLIDE-owned application instance.',
                eventIndex: entry.index,
            });
        }
        if (entry.index <= openIndex || (firstCloseOrKillIndex >= 0 && entry.index > firstCloseOrKillIndex)) {
            issues.push({
                code: 'macro-order',
                message: 'VBA test macros must run after the file opens and before close or kill cleanup.',
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
        if (hostId && entry.event.hostId !== hostId) {
            issues.push({
                code: 'macro-instance',
                message: 'Every VBA test macro result must come from the single XLIDE-owned application instance.',
                eventIndex: entry.index,
            });
        }
    }
    for (const entry of modalBlocked) {
        const resultAfterModal = macroFinished.find((finished) =>
            finished.index > entry.index &&
            finished.event.hostId === entry.event.hostId &&
            finished.event.qualifiedName === entry.event.qualifiedName &&
            finished.event.outcome === 'modal-blocked',
        );
        if (!resultAfterModal) {
            issues.push({
                code: 'modal-result',
                message: 'A blocked modal dialog must be reflected as a modal-blocked macro result.',
                eventIndex: entry.index,
            });
            continue;
        }
        const killAfterModal = killed.find((kill) =>
            kill.index > resultAfterModal.index && kill.event.hostId === entry.event.hostId,
        );
        if (!killAfterModal) {
            issues.push({
                code: 'modal-cleanup',
                message: 'A blocked modal dialog must clean up the XLIDE-owned application instance.',
                eventIndex: resultAfterModal.index,
            });
        }
    }

    const firstHang = macroFinished.find((entry) =>
        entry.event.outcome === 'timeout' || entry.event.outcome === 'hung' || entry.event.outcome === 'modal-blocked',
    );
    if (firstHang) {
        const killAfterHang = killed.find((entry) =>
            entry.event.hostId === firstHang.event.hostId && entry.index > firstHang.index,
        );
        if (!killAfterHang) {
            issues.push({
                code: 'hang-cleanup',
                message: 'A timeout or hang must clean up the XLIDE-owned application instance.',
                eventIndex: firstHang.index,
            });
        }
        const macroAfterKill = killAfterHang
            ? macroStarted.find((entry) => entry.index > killAfterHang.index)
            : undefined;
        if (macroAfterKill) {
            issues.push({
                code: 'no-macros-after-kill',
                message: 'No further VBA test macros may run after the owned application instance is killed.',
                eventIndex: macroAfterKill.index,
            });
        }
        return issues;
    }

    const closeEntry = closed[0];
    if (!closeEntry || closeEntry.event.hostId !== hostId || closeEntry.event.saveChanges) {
        issues.push({
            code: 'close-without-saving',
            message: 'Normal VBA test runs must close the file without saving changes.',
            eventIndex: closeEntry?.index,
        });
    }
    const quitAfterClose = closeEntry
        ? quit.find((entry) => entry.event.hostId === closeEntry.event.hostId && entry.index > closeEntry.index)
        : undefined;
    if (!quitAfterClose) {
        issues.push({
            code: 'normal-cleanup',
            message: 'Normal VBA test runs must quit the XLIDE-owned application instance after closing the file.',
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
