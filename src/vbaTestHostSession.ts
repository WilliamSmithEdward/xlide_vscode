import * as cp from 'child_process';
import { DEFAULT_VBA_TEST_TIMEOUT_MS } from './vbaTestExcelHost';
import {
    parseVbaTestHostEventLine,
    validateVbaTestHostOracleTrace,
    type VbaTestHostOracleEvent,
} from './vbaTestHostOracle';
import { vbaTestFailureMessage } from './vbaTestFailureMessages';
import { errorMessage } from './util/errors';
import {
    runPowerShell,
    type PowerShellRun,
    type RunPowerShellResult,
} from './util/powershell';

/**
 * The owned read-only Excel test host session: drives one staged
 * run-vba-tests.ps1 child process and aggregates its oracle event stream into
 * per-test results (formerly a 400-line Promise executor with 15 closure
 * variables in vbaTestExecution.ts).
 *
 * Lifecycle phases:
 *  - startup: from spawn until the first macro-started event, guarded by the
 *    startup watchdog.
 *  - macro: one currentMacro at a time, guarded by the per-macro watchdog and
 *    interrupted when the host reports an unsafe modal dialog (modal-blocked).
 *  - cleanup: after workbook-closed / excel-quit, guarded by the cleanup
 *    watchdog while the host script finishes COM release.
 *  - settled: finish() resolved the session promise (exactly once).
 *
 * abort() is the single failure path shared by the startup watchdog, the
 * macro watchdog, the modal-blocked handler, and the host spawn-error / exit
 * handlers: it synthesizes a terminal macro-finished event, kills the owned
 * Excel process, optionally kills the PowerShell host, and settles.
 */

export interface OwnedReadOnlyExcelHostRunResult {
    events: VbaTestHostOracleEvent[];
    resultsByName: Map<string, OwnedReadOnlyExcelHostTestResult>;
    hostError?: string;
    timedOutAfter?: string;
}

export interface OwnedReadOnlyExcelHostTestResult {
    outcome: 'passed' | 'failed' | 'timeout' | 'modal-blocked' | 'runner-error';
    durationMs: number;
    message?: string;
    errorNumber?: number;
    errorSource?: string;
    output?: string[];
}

export interface OwnedExcelTestHostSessionOptions {
    hostScriptPath: string;
    /** Disposes the staged temp dir once the session settles. */
    disposeStaging: () => void;
    log: (message: string) => void;
}

type OwnedExcelKillReason = 'timeout' | 'hung' | 'modal-blocked' | 'runner-error' | 'cleanup-failed';

const DEFAULT_VBA_TEST_CLEANUP_GRACE_MS = 5000;

type MacroStartedEvent = Extract<VbaTestHostOracleEvent, { kind: 'macro-started' }>;
type MacroFinishedEvent = Extract<VbaTestHostOracleEvent, { kind: 'macro-finished' }>;
type ModalBlockedEvent = Extract<VbaTestHostOracleEvent, { kind: 'modal-blocked' }>;

interface CurrentMacro {
    excelId: string;
    qualifiedName: string;
    timeoutMs: number;
    startedMs: number;
}

interface AbortOptions {
    reason: OwnedExcelKillReason;
    hostError: string;
    /** Terminal macro-finished event pushed before killing; omitted when no macro is in flight. */
    finishedEvent?: MacroFinishedEvent;
    /** Also kill the PowerShell host process (it already exited in the spawn-error/exit paths). */
    killHostProcess?: boolean;
}

export function runOwnedExcelTestHostSession(
    options: OwnedExcelTestHostSessionOptions,
): Promise<OwnedReadOnlyExcelHostRunResult> {
    return new Promise<OwnedReadOnlyExcelHostRunResult>((resolve) => {
        // The session keeps itself alive through the watchdog timers and
        // child-process callbacks it registers.
        void new OwnedExcelTestHostSession(options, resolve);
    });
}

class OwnedExcelTestHostSession {
    private readonly events: VbaTestHostOracleEvent[] = [];
    private currentMacro: CurrentMacro | undefined;
    private currentModalBlocker: ModalBlockedEvent | undefined;
    private startupWatchdog: ReturnType<typeof setTimeout> | undefined;
    private macroWatchdog: ReturnType<typeof setTimeout> | undefined;
    private cleanupWatchdog: ReturnType<typeof setTimeout> | undefined;
    private ownedExcelPid: number | undefined;
    private ownedExcelKilled = false;
    private sawWorkbookClosed = false;
    private sawExcelQuit = false;
    private timedOutAfter: string | undefined;
    private settled = false;

    private readonly log: (message: string) => void;
    private readonly disposeStaging: () => void;
    private readonly hostRun: PowerShellRun;

    constructor(
        options: OwnedExcelTestHostSessionOptions,
        private readonly resolve: (result: OwnedReadOnlyExcelHostRunResult) => void,
    ) {
        this.log = options.log;
        this.disposeStaging = options.disposeStaging;
        this.armStartupWatchdog();
        this.hostRun = runPowerShell({
            args: ['-File', options.hostScriptPath],
            onSpawn: (pid) => {
                this.log(`[runVbaTests] Spawned owned host powershell.exe (pid=${pid ?? 'unknown'})`);
            },
            onStdoutLine: (line) => this.handleStdoutLine(line),
            onStderrLine: (line) => {
                this.log(`[runVbaTests host stderr] ${line}`);
            },
        });
        void this.hostRun.result.then((result) => this.handleHostExit(result));
    }

    // ----- phase: startup ---------------------------------------------------

    private armStartupWatchdog(): void {
        this.startupWatchdog = setTimeout(() => {
            if (this.settled || this.currentMacro) {
                return;
            }
            this.timedOutAfter = 'test host startup';
            const excelId = this.events.find((event) => event.kind === 'excel-created')?.excelId ?? 'unknown';
            const message = `Excel test host timed out before starting tests after ${DEFAULT_VBA_TEST_TIMEOUT_MS} ms.`;
            this.abort({
                reason: 'timeout',
                hostError: message,
                killHostProcess: true,
                finishedEvent: {
                    kind: 'macro-finished',
                    excelId,
                    qualifiedName: 'XLIDE.TestHostStartup',
                    outcome: 'timeout',
                    durationMs: DEFAULT_VBA_TEST_TIMEOUT_MS,
                    message,
                },
            });
        }, DEFAULT_VBA_TEST_TIMEOUT_MS);
    }

    // ----- phase: macro -----------------------------------------------------

    private armMacroWatchdog(event: MacroStartedEvent): void {
        this.clearStartupWatchdog();
        this.clearMacroWatchdog();
        const timeoutMs = event.timeoutMs ?? DEFAULT_VBA_TEST_TIMEOUT_MS;
        this.currentMacro = {
            excelId: event.excelId,
            qualifiedName: event.qualifiedName,
            timeoutMs,
            startedMs: Date.now(),
        };
        this.macroWatchdog = setTimeout(() => {
            const macro = this.currentMacro;
            if (!macro || this.timedOutAfter) {
                return;
            }
            this.timedOutAfter = macro.qualifiedName;
            const durationMs = Date.now() - macro.startedMs;
            // Note (audit #39): currentModalBlocker is only assigned by the
            // modal-blocked handler, which aborts the session synchronously,
            // so the modal branch below is not reachable in any observed
            // trace; kept verbatim from the pre-decomposition state machine.
            const modalBlocker = this.currentModalBlocker;
            const modalDetail = modalBlocker
                ? [modalBlocker.title, modalBlocker.message].filter(Boolean).join(': ')
                : '';
            const outcome = modalBlocker ? 'modal-blocked' : 'timeout';
            const message = modalBlocker
                ? `Blocked by Excel modal dialog${modalDetail ? ` (${modalDetail})` : ''}.`
                : `Timed out after ${macro.timeoutMs} ms.`;
            this.abort({
                reason: modalBlocker ? 'modal-blocked' : 'timeout',
                hostError: message,
                killHostProcess: true,
                finishedEvent: {
                    kind: 'macro-finished',
                    excelId: macro.excelId,
                    qualifiedName: macro.qualifiedName,
                    outcome,
                    durationMs,
                    message,
                },
            });
        }, timeoutMs);
    }

    // ----- phase: cleanup ---------------------------------------------------

    private armCleanupWatchdog(stage: 'workbook-closed' | 'excel-quit'): void {
        this.clearCleanupWatchdog();
        this.cleanupWatchdog = setTimeout(() => {
            if (this.settled) {
                return;
            }
            const elapsedDescription = `${DEFAULT_VBA_TEST_CLEANUP_GRACE_MS} ms`;
            if (!this.sawExcelQuit) {
                this.log(`[runVbaTests] Cleanup watchdog elapsed ${elapsedDescription} after workbook close; killing owned Excel.`);
                this.killOwnedExcel('cleanup-failed');
            } else {
                this.log(`[runVbaTests] Cleanup watchdog elapsed ${elapsedDescription} after Excel quit; stopping host script.`);
            }
            this.hostRun.kill();
            this.finish();
        }, DEFAULT_VBA_TEST_CLEANUP_GRACE_MS);
        this.log(`[runVbaTests] Cleanup watchdog armed after ${stage} (${DEFAULT_VBA_TEST_CLEANUP_GRACE_MS} ms).`);
    }

    // ----- event stream -----------------------------------------------------

    private handleStdoutLine(line: string): void {
        const trimmed = line.trim();
        if (!trimmed) {
            return;
        }
        try {
            const event = parseVbaTestHostEventLine(trimmed);
            if (event) {
                this.handleEvent(event);
            } else {
                this.log(`[runVbaTests host stdout] ${trimmed}`);
            }
        } catch (err) {
            this.log(`[runVbaTests host stdout] ${trimmed}`);
            this.log(`[runVbaTests host parse] ${errorMessage(err)}`);
        }
    }

    private handleEvent(event: VbaTestHostOracleEvent): void {
        this.events.push(event);
        if (event.kind === 'excel-created') {
            this.ownedExcelPid = event.pid;
            this.log(`[runVbaTests host] excel-created pid=${this.ownedExcelPid ?? 'unknown'}`);
        } else if (event.kind === 'host-phase') {
            this.log(`[runVbaTests host] phase ${event.phase} ${event.outcome} (${event.durationMs} ms)`);
        } else if (event.kind === 'macro-started') {
            this.log(`[runVbaTests host] macro-started ${event.qualifiedName} timeoutMs=${event.timeoutMs ?? DEFAULT_VBA_TEST_TIMEOUT_MS}`);
            this.currentModalBlocker = undefined;
            this.armMacroWatchdog(event);
        } else if (event.kind === 'modal-detected') {
            this.log(`[runVbaTests host] modal-detected ${event.qualifiedName} classification=${event.classification ?? 'unknown'} safeToDismiss=${event.safeToDismiss ?? false}`);
        } else if (event.kind === 'modal-dismissed') {
            this.log(`[runVbaTests host] modal-dismissed ${event.qualifiedName} button=${event.button ?? 'unknown'} dismissed=${event.dismissed}`);
        } else if (event.kind === 'modal-blocked') {
            if (this.currentMacro &&
                this.currentMacro.excelId === event.excelId &&
                this.currentMacro.qualifiedName === event.qualifiedName) {
                const macro = this.currentMacro;
                this.currentModalBlocker = event;
                const durationMs = Date.now() - macro.startedMs;
                const message = modalBlockedMessage(event);
                this.abort({
                    reason: 'modal-blocked',
                    hostError: message,
                    killHostProcess: true,
                    finishedEvent: {
                        kind: 'macro-finished',
                        excelId: macro.excelId,
                        qualifiedName: macro.qualifiedName,
                        outcome: 'modal-blocked',
                        durationMs,
                        message,
                    },
                });
            }
            this.log(`[runVbaTests host] modal-blocked ${event.qualifiedName}: ${event.reason}`);
        } else if (event.kind === 'macro-finished') {
            this.log(`[runVbaTests host] macro-finished ${event.qualifiedName} outcome=${event.outcome}`);
            this.clearMacroWatchdog();
            this.currentMacro = undefined;
            this.currentModalBlocker = undefined;
        } else if (event.kind === 'workbook-closed') {
            this.sawWorkbookClosed = true;
            this.log(`[runVbaTests host] workbook-closed durationMs=${event.durationMs ?? 'unknown'}`);
            this.armCleanupWatchdog('workbook-closed');
        } else if (event.kind === 'excel-quit') {
            this.sawExcelQuit = true;
            this.log(`[runVbaTests host] excel-quit durationMs=${event.durationMs ?? 'unknown'}`);
            // Arm unconditionally. If $workbook.Close($false) threw, the host
            // emits no workbook-closed event but still proceeds to $excel.Quit(),
            // leaving the COM-release phase (WaitForPendingFinalizers, which can
            // hang on a stuck COM object) with NO active watchdog and the whole
            // run hanging forever. Always guarding cleanup prevents that.
            this.armCleanupWatchdog('excel-quit');
        }
    }

    private handleHostExit(result: RunPowerShellResult): void {
        if (result.spawnError) {
            const message = `RUNNER_FAILED|${result.spawnError.message}`;
            this.abort({
                reason: 'runner-error',
                hostError: message,
                finishedEvent: this.currentMacroFinishedEvent('runner-error', message),
            });
            return;
        }
        this.log(`[runVbaTests] owned host powershell exited with code=${result.code} signal=${result.signal ?? 'none'}`);
        if (this.settled) {
            return;
        }
        if (result.code === 0) {
            this.finish();
            return;
        }
        const sentinel = result.stderrLines.find((line) => line.includes('XLIDE_TEST_HOST_ERROR|'));
        const hostError = sentinel
            ? sentinel.slice(sentinel.indexOf('XLIDE_TEST_HOST_ERROR|') + 'XLIDE_TEST_HOST_ERROR|'.length)
            : result.stderrLines.join('\n') || `PowerShell exited with code ${result.code}`;
        this.abort({
            reason: 'runner-error',
            hostError,
            finishedEvent: this.currentMacroFinishedEvent('runner-error', hostError),
        });
    }

    // ----- terminal transitions ----------------------------------------------

    /**
     * Single failure path: push the synthesized terminal event (if any), kill
     * the owned Excel process, optionally kill the PowerShell host, settle.
     */
    private abort(options: AbortOptions): void {
        if (options.finishedEvent) {
            this.events.push(options.finishedEvent);
        }
        this.killOwnedExcel(options.reason);
        if (options.killHostProcess) {
            this.hostRun.kill();
        }
        this.finish(options.hostError);
    }

    private finish(hostError?: string): void {
        if (this.settled) {
            return;
        }
        this.settled = true;
        this.clearAllWatchdogs();
        this.disposeStaging();
        const resultsByName = new Map<string, OwnedReadOnlyExcelHostTestResult>();
        for (const event of this.events) {
            const result = this.eventResult(event);
            if (result && event.kind === 'macro-finished') {
                resultsByName.set(event.qualifiedName, result);
            }
        }
        const oracleIssues = validateVbaTestHostOracleTrace(this.events);
        for (const issue of oracleIssues) {
            this.log(`[runVbaTests oracle] ${issue.code}: ${issue.message}`);
        }
        this.resolve({
            events: this.events,
            resultsByName,
            hostError: hostError === undefined ? undefined : vbaTestFailureMessage(hostError),
            timedOutAfter: this.timedOutAfter,
        });
    }

    private killOwnedExcel(reason: OwnedExcelKillReason): void {
        if (!this.ownedExcelPid || this.ownedExcelKilled) {
            return;
        }
        this.ownedExcelKilled = true;
        this.log(`[runVbaTests] Killing owned Excel process ${this.ownedExcelPid} after ${reason}.`);
        // taskkill runs on the abort/cleanup recovery paths; a spawn failure
        // (EMFILE/ENOMEM under load, or a locked-down PATH) must be logged, not
        // left to surface as an uncaught child 'error' that crashes the host.
        const killer = cp.spawn('taskkill.exe', ['/PID', String(this.ownedExcelPid), '/T', '/F'], { windowsHide: true });
        killer.on('error', (err) => {
            this.log(`[runVbaTests] taskkill failed for pid ${this.ownedExcelPid}: ${errorMessage(err)}`);
        });
        const excelId = this.currentMacro?.excelId ?? this.events.find((event) => event.kind === 'excel-created')?.excelId;
        if (excelId) {
            this.events.push({ kind: 'excel-killed', excelId, reason });
        }
    }

    // ----- helpers ------------------------------------------------------------

    /** Synthesizes a terminal macro-finished event for the in-flight macro, if any. */
    private currentMacroFinishedEvent(
        outcome: 'runner-error',
        message: string,
    ): MacroFinishedEvent | undefined {
        const macro = this.currentMacro;
        if (!macro) {
            return undefined;
        }
        return {
            kind: 'macro-finished',
            excelId: macro.excelId,
            qualifiedName: macro.qualifiedName,
            outcome,
            durationMs: Date.now() - macro.startedMs,
            message,
        };
    }

    private eventResult(event: VbaTestHostOracleEvent): OwnedReadOnlyExcelHostTestResult | undefined {
        if (event.kind !== 'macro-finished') {
            return undefined;
        }
        if (event.outcome === 'passed' || event.outcome === 'failed') {
            return {
                outcome: event.outcome,
                durationMs: event.durationMs ?? 0,
                message: event.message,
                errorNumber: event.errorNumber,
                errorSource: event.errorSource,
                output: event.output,
            };
        }
        if (event.outcome === 'timeout' || event.outcome === 'hung') {
            return {
                outcome: 'timeout',
                durationMs: event.durationMs ?? this.currentMacro?.timeoutMs ?? DEFAULT_VBA_TEST_TIMEOUT_MS,
                message: event.message,
            };
        }
        if (event.outcome === 'modal-blocked') {
            return {
                outcome: 'modal-blocked',
                durationMs: event.durationMs ?? this.currentMacro?.timeoutMs ?? DEFAULT_VBA_TEST_TIMEOUT_MS,
                message: event.message,
            };
        }
        return {
            outcome: 'runner-error',
            durationMs: event.durationMs ?? 0,
            message: event.message,
        };
    }

    private clearStartupWatchdog(): void {
        if (this.startupWatchdog) {
            clearTimeout(this.startupWatchdog);
            this.startupWatchdog = undefined;
        }
    }

    private clearMacroWatchdog(): void {
        if (this.macroWatchdog) {
            clearTimeout(this.macroWatchdog);
            this.macroWatchdog = undefined;
        }
    }

    private clearCleanupWatchdog(): void {
        if (this.cleanupWatchdog) {
            clearTimeout(this.cleanupWatchdog);
            this.cleanupWatchdog = undefined;
        }
    }

    /** Single watchdog teardown used when the session settles. */
    private clearAllWatchdogs(): void {
        this.clearMacroWatchdog();
        this.clearStartupWatchdog();
        this.clearCleanupWatchdog();
    }
}

function modalBlockedMessage(event: ModalBlockedEvent): string {
    const modalDetail = [event.title, event.message].filter(Boolean).join(': ');
    return `Blocked by Excel modal dialog${modalDetail ? ` (${modalDetail})` : ''}.`;
}
