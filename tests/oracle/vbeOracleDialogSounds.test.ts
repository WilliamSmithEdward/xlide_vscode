import { spawnSync as realSpawnSync } from 'node:child_process';
import { EventEmitter } from 'node:events';
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, win32 } from 'node:path';
import readline from 'node:readline';
import { PassThrough } from 'node:stream';
import { runInNewContext } from 'node:vm';
import { describe, expect, it, vi } from 'vitest';
import { OFFICE_HOST_APPS } from '../../src/officeHostApps';

// Run the coordinator's actual functions, with no real PowerShell, signals,
// registry writes, Office processes or account-profile writes.
const source = readFileSync(join(process.cwd(), 'syntax_corpus/oracle/run_vbe_oracle.mjs'), 'utf8');
const names = ['dialogSoundsPath', 'dialogSoundsScript', 'setDialogSounds', 'parseArgs',
    'caseHost', 'unsupportedHost', 'malformedReferences', 'hostsNeedingBuilds', 'knownHostCrashResult', 'expectedMatches', 'isOracleInfrastructureFailure', 'isObserveWithoutVerdict'];
const functions = names.map((name) => {
    const start = source.indexOf(`function ${name}(`);
    // Up to the next top-level declaration, so a closing brace at the start of
    // a line inside a template literal cannot cut a function short.
    const next = source.slice(start + 1).search(/\n(?:async function |function |const |let |\/\/ )/);
    return next < 0 ? source.slice(start) : source.slice(start, start + 1 + next);
}).join('\n');
// The worker's sound function, command and reports, taken the same way.
const workerSounds = ['dialogSoundsFunction', 'dialogSoundsEnv', 'reportDialogSounds', 'reportWorkerSounds', 'workerCommand'].map((name) => {
    const start = source.indexOf(`function ${name}(`);
    const next = source.slice(start + 1).search(/\n(?:async function |function |const |let |\/\/ )/);
    return next < 0 ? source.slice(start) : source.slice(start, start + 1 + next);
}).join('\n');
const soundsConstant = source.match(/^const DIALOG_SOUNDS = .*;$/m)![0];
const prefixConstant = source.match(/^const SOUNDS_PREFIX = .*;$/m)![0];
const mainStart = source.indexOf('async function main(');
const mainSource = source.slice(mainStart, source.indexOf('\nprocess.exit(', mainStart));
const signals = [['SIGINT', 130], ['SIGBREAK', 149], ['SIGTERM', 143], ['SIGHUP', 129]] as const;
const ownedEvents = ['exit', ...signals.map(([signal]) => signal)];
type Case = { id: string; host?: string; expected?: string; crashesHostBuilds?: string[] };
type Completed = { status: number | null; stderr?: string | null; error?: Error };
type SpawnOptions = { encoding: string; timeout: number; env: Record<string, string> };
const soundsFile = 'C:\\Users\\account\\.xlide-vbe-oracle-sounds.json';

function deferred<T>() {
    let resolve!: (value: T) => void;
    let reject!: (reason: Error) => void;
    const promise = new Promise<T>((yes, no) => { resolve = yes; reject = no; });
    return { promise, resolve, reject };
}

/** The run lock is taken asynchronously (issue #1438): let main reach its first case. */
const lockTaken = (): Promise<void> => new Promise((resolve) => setImmediate(resolve));

// soundsLeft: the saved-sounds file is still there when the coordinator looks,
// as when its worker was killed or started none.
function setup(cases: Case[] = [{ id: 'first' }, { id: 'second' }], lockHeld = true, soundsLeft = true) {
    const trace: string[] = [];
    const workers = { started: 0 };
    const hostSweeps = { count: 0 };
    const errors: string[] = [];
    const output: string[] = [];
    const interrupted = new Error('fake process exited');
    const fakeProcess = Object.assign(new EventEmitter(), {
        platform: 'win32', pid: 4321,
        env: { PRESERVED: 'yes', USERPROFILE: 'C:\\overridden-profile',
            XLIDE_SOUNDS_FILE: 'wrong', XLIDE_SOUNDS_OWNER: 'wrong', XLIDE_SOUNDS_ACTION: 'wrong' },
        exit: vi.fn((code: number): never => {
            trace.push(`exit:${code}`);
            fakeProcess.emit('exit', code);
            // Real process.exit never returns. The suspended main is unwound
            // separately by rejecting its fake case/close promise below.
            throw interrupted;
        }),
    });
    const os = {
        userInfo: vi.fn(() => ({ homedir: 'C:\\Users\\account' })),
        homedir: vi.fn(() => 'C:\\Users\\fallback'),
    };
    const spawnSync = vi.fn((_command: string, _args: string[], _options: SpawnOptions): Completed => {
        trace.push('restore');
        return { status: 0 };
    });
    const existsSync = vi.fn((_file: string) => soundsLeft);
    const close = vi.fn(async () => {});
    const kill = vi.fn(() => { trace.push('kill'); });
    const runCaseWithRetries = vi.fn(async (_worker: unknown, kase: Case) => ({ caseId: kase.id, outcome: 'accepted' }));
    const comRegistered = vi.fn(() => true);
    const registeredHostBuild = vi.fn((): { registered: boolean; build: string | null; why?: string } => ({
        registered: true, build: '16.0.20430.20182',
    }));
    const loadCasesDocument = vi.fn(() => ({ cases }));
    const api = runInNewContext(`
        const DEFAULT_HOST = 'excel';
        const DEFAULT_CASES = 'cases.json';
        const DEFAULT_TIMEOUT_RETRIES = 2;
        const DEFAULT_CASES_PER_INSTANCE = 40;
        const USAGE = 'usage';
        const EVIDENCE_OUTCOMES = new Set(['accepted', 'rejected']);
        ${soundsConstant}
        ${prefixConstant}
        ${functions}
        ${workerSounds}
        ${mainSource}
        ({ main, dialogSoundsPath, setDialogSounds, reportWorkerSounds, workerCommand })`, {
        process: fakeProcess, os, path: win32, spawnSync,
        console: { log: (line: string) => output.push(line), error: (line: string) => errors.push(line) },
        OFFICE_HOSTS: Object.fromEntries(Object.entries(OFFICE_HOST_APPS).map(([host, info]) => [host,
            { progId: info.progId, image: info.processName, noun: info.noun }])),
        loadCasesDocument, comRegistered,
        registeredHostBuild,
        loadCheckpoint: () => new Map(),
        // The run lock is covered in vbeOracleRunLock.test.ts; here it is always free.
        acquireOracleRunLock: async () => (lockHeld
            ? { held: true, release: async () => {} }
            : { held: false, reason: 'busy', owner: { pid: 99 } }),
        oracleRunLockAddress: () => 'test-lock', oracleRunOwner: () => ({ pid: 1 }), describeRefusal: () => 'refused',
        // Owned hosts are covered in vbeOracleOwnedHosts.test.ts; here the run owns none.
        OwnedHosts: class { load() { return []; } add() {} sweep() { hostSweeps.count += 1; return []; } },
        ownedHostsPath: () => 'hosts.json', describeLeftHosts: () => 'left', abandonActiveCases: () => {},
        unwritableCheckpoint: () => undefined, // A real path is checked in vbeOracleOptions.test.ts.
        OracleWorker: class {
            constructor() { workers.started++; }
            async close() {
                trace.push('close:begin');
                try { await close(); } finally { trace.push('close:end'); }
            }
            kill() { kill(); }
        },
        caseHash: () => 'hash', CHECKPOINT_KEY_VERSION: 2, checkpointKey: (id: string) => id, phaseMatches: () => true,
        runCaseWithRetries: (worker: unknown, kase: Case) => {
            trace.push(`measure:${kase.id}`);
            return runCaseWithRetries(worker, kase);
        },
        fs: { appendFileSync: vi.fn(), existsSync },
    }) as {
        main: (argv: string[]) => Promise<number>;
        dialogSoundsPath: () => string;
        setDialogSounds: (action: string) => void;
        reportWorkerSounds: (text: string) => void;
        workerCommand: (casesPerInstance: number) => string;
    };
    return { ...api, trace, workers, hostSweeps, errors, output, fakeProcess, os, spawnSync, existsSync, close, kill,
        runCaseWithRetries, comRegistered, registeredHostBuild, loadCasesDocument, interrupted };
}

function expectNoOwnedListeners(probe: ReturnType<typeof setup>) {
    for (const event of ownedEvents) { expect(probe.fakeProcess.listenerCount(event), event).toBe(0); }
}

describe('VBE oracle dialog sound lifecycle', () => {
    it('starts no worker and leaves the sounds alone while another run holds the lock', async () => {
        const probe = setup([{ id: 'first' }], false);
        expect(await probe.main(['--json'])).toBe(3);
        expect(probe.trace).toEqual([]);
        expect(probe.workers.started).toBe(0);
        expect(probe.errors).toEqual(['refused']);
        expect(probe.spawnSync).not.toHaveBeenCalled();
    });

    // The worker's process silences the sounds before its first case
    // ('worker command' below), so the coordinator starts nothing before it.
    it.each(Object.keys(OFFICE_HOST_APPS))('puts back sounds left after %s closes, once, and removes its listeners', async (host) => {
        const probe = setup([{ id: 'first', host }, { id: 'second', host }]);
        const unrelated = vi.fn();
        for (const event of ownedEvents) { probe.fakeProcess.on(event, unrelated); }
        expect(await probe.main(['--json'])).toBe(0);
        expect(probe.trace).toEqual(['measure:first', 'measure:second', 'close:begin', 'close:end', 'restore']);
        expect(probe.existsSync).toHaveBeenCalledExactlyOnceWith(soundsFile);
        expect(probe.errors).toEqual([]);
        for (const event of ownedEvents) { expect(probe.fakeProcess.listeners(event)).toEqual([unrelated]); }
        probe.fakeProcess.emit('exit', 0);
        expect(probe.spawnSync).toHaveBeenCalledOnce();
    });

    it.each(Object.keys(OFFICE_HOST_APPS))('starts no PowerShell for the sounds when the %s worker put them back', async (host) => {
        const probe = setup([{ id: 'first', host }], true, false);
        expect(await probe.main(['--json'])).toBe(0);
        expect(probe.trace).toEqual(['measure:first', 'close:begin', 'close:end']);
        expect(probe.existsSync).toHaveBeenCalledExactlyOnceWith(soundsFile);
        expect(probe.spawnSync).not.toHaveBeenCalled();
    });

    it('puts back a killed run\'s sounds when every case is skipped and no worker starts', async () => {
        const probe = setup();
        probe.registeredHostBuild.mockReturnValue({ registered: true, build: '16.0.1.2' });
        probe.loadCasesDocument.mockReturnValue({ cases: [{ id: 'skipped', crashesHostBuilds: ['16.0.1.2'] }] });
        expect(await probe.main(['--json'])).toBe(0);
        expect(probe.trace).toEqual(['close:begin', 'close:end', 'restore']);
        expect(probe.runCaseWithRetries).not.toHaveBeenCalled();
    });

    it.each(['case', 'close', 'both'])('restores and removes listeners when %s throws', async (failure) => {
        const probe = setup();
        const caseError = new Error('case failed');
        const closeError = new Error('close failed');
        if (failure !== 'close') { probe.runCaseWithRetries.mockRejectedValueOnce(caseError); }
        if (failure !== 'case') { probe.close.mockRejectedValueOnce(closeError); }
        await expect(probe.main([])).rejects.toBe(failure === 'case' ? caseError : closeError);
        expect(probe.trace).toEqual(['measure:first',
            ...(failure === 'close' ? ['measure:second'] : []), 'close:begin', 'close:end', 'restore']);
        expect(probe.close).toHaveBeenCalledOnce();
        expectNoOwnedListeners(probe);
        probe.fakeProcess.emit('exit', 1);
        expect(probe.spawnSync).toHaveBeenCalledOnce();
    });

    it('cleans up synchronously on exit, without restoring again in finally', async () => {
        const probe = setup([{ id: 'first' }]);
        const pending = deferred<{ caseId: string; outcome: string }>();
        probe.runCaseWithRetries.mockReturnValueOnce(pending.promise);
        const running = probe.main(['--json']);
        await lockTaken();
        expect(probe.trace).toEqual(['measure:first']);
        probe.fakeProcess.emit('exit', 1);
        // The worker and the hosts it started are ended, then the sounds put back (issue #1439).
        expect(probe.trace).toEqual(['measure:first', 'kill', 'restore']);
        expect(probe.hostSweeps.count).toBeGreaterThan(0);
        probe.fakeProcess.emit('exit', 1);
        pending.resolve({ caseId: 'first', outcome: 'accepted' });
        expect(await running).toBe(0);
        expect(probe.trace).toEqual(['measure:first', 'kill', 'restore', 'close:begin', 'close:end']);
        expectNoOwnedListeners(probe);
    });

    it.each(signals.flatMap(([signal, code]) => ['case', 'close'].map((stage) => ({ signal, code, stage }))))(
        'ends the worker and its hosts, restores and exits $code on $signal during $stage, without a second restore', async ({ signal, code, stage }) => {
            const probe = setup([{ id: 'first' }]);
            const measuring = deferred<{ caseId: string; outcome: string }>();
            const closing = deferred<void>();
            const closeStarted = deferred<void>();
            probe.runCaseWithRetries.mockReturnValueOnce(measuring.promise);
            probe.close.mockImplementationOnce(() => { closeStarted.resolve(); return closing.promise; });
            const running = probe.main(['--json']);
            const rejected = expect(running).rejects.toBe(probe.interrupted);
            await lockTaken();
            if (stage === 'close') {
                measuring.resolve({ caseId: 'first', outcome: 'accepted' });
                await closeStarted.promise;
            }
            expect(() => probe.fakeProcess.emit(signal)).toThrow(probe.interrupted);
            expect(probe.trace.slice(-3)).toEqual(['kill', 'restore', `exit:${code}`]);
            expect(probe.fakeProcess.exit).toHaveBeenCalledExactlyOnceWith(code);
            // The coordinator ends the worker and the hosts it recorded itself (issue
            // #1439), whether or not the worker also saw the Ctrl-C.
            expect(probe.kill).toHaveBeenCalledTimes(1);
            expect(probe.hostSweeps.count).toBeGreaterThan(0);
            expect(probe.fakeProcess.listenerCount(signal)).toBe(0);
            if (stage === 'case') {
                measuring.reject(probe.interrupted);
                closing.resolve();
            } else { closing.reject(probe.interrupted); }
            await rejected;
            expect(probe.trace.filter((event) => event === 'restore')).toHaveLength(1);
            expect(probe.spawnSync).toHaveBeenCalledOnce();
            expectNoOwnedListeners(probe);
        });
    it.each([
        { argv: ['--invalid'], platform: 'win32', code: 2 },
        { argv: ['--timeout', 'oops'], platform: 'win32', code: 2 },
        { argv: ['--timeout', '0'], platform: 'win32', code: 2 },
        { argv: ['--timeout-retries', '-1'], platform: 'win32', code: 2 },
        { argv: ['--dialog-hold-seconds', '-1'], platform: 'win32', code: 2 },
        { argv: ['--cases-per-instance', '0'], platform: 'win32', code: 2 },
        { argv: ['--promote-observed'], platform: 'win32', code: 2 },
        { argv: ['--case', 'missing'], platform: 'win32', code: 2 },
        // An option with no value, or another option in its place (issue #1440).
        { argv: ['--checkpoint'], platform: 'win32', code: 2 },
        { argv: ['--checkpoint', '--strict'], platform: 'win32', code: 2 },
        { argv: ['--cases', '--json'], platform: 'win32', code: 2 },
        { argv: ['--case'], platform: 'win32', code: 2 },
        { argv: ['--timeout', '5abc'], platform: 'win32', code: 2 },
        { argv: [], platform: 'linux', code: 2 },
        { argv: [], platform: 'darwin', code: 2 },
        { argv: ['--help'], platform: 'win32', code: 0 },
    ])('does not touch sounds for $argv on $platform', async ({ argv, platform, code }) => {
        const probe = setup();
        probe.fakeProcess.platform = platform;
        expect(await probe.main(argv)).toBe(code);
        expect(probe.spawnSync).not.toHaveBeenCalled();
        expect(probe.existsSync).not.toHaveBeenCalled();
        expect(probe.runCaseWithRetries).not.toHaveBeenCalled();
        expectNoOwnedListeners(probe);
    });

    it.each(['unknown host', 'capitalized host', 'null host', 'unregistered host', 'invalid document', 'unreadable checkpoint build'])(
        'does not touch sounds after startup failure: %s', async (failure) => {
            const hosts: Record<string, string | null> = { 'unknown host': 'other', 'capitalized host': 'Excel', 'null host': null };
            const host = failure in hosts ? hosts[failure] : 'excel';
            const probe = setup([{ id: 'first', host: host as string }]);
            if (failure === 'unregistered host') { probe.comRegistered.mockReturnValue(false); }
            if (failure === 'invalid document') { probe.loadCasesDocument.mockImplementation(() => { throw new Error('bad JSON'); }); }
            if (failure === 'unreadable checkpoint build') {
                probe.registeredHostBuild.mockReturnValue({ registered: true, build: null, why: 'unreadable' });
            }
            expect(await probe.main(failure === 'unreadable checkpoint build' ? ['--checkpoint', 'checkpoint'] : [])).toBe(2);
            expect(probe.spawnSync).not.toHaveBeenCalled();
            expect(probe.existsSync).not.toHaveBeenCalled();
            expectNoOwnedListeners(probe);
        });
});

describe('VBE oracle dialog sound PowerShell invocation', () => {
    it('passes the account-profile path, pid, action and inherited environment to the restore', () => {
        const probe = setup();
        probe.os.userInfo.mockReturnValue({ homedir: "C:\\Users\\O'Brien [account]" });
        probe.setDialogSounds('restore');
        const file = "C:\\Users\\O'Brien [account]\\.xlide-vbe-oracle-sounds.json";
        expect(probe.existsSync).toHaveBeenCalledExactlyOnceWith(file);
        expect(probe.spawnSync).toHaveBeenCalledExactlyOnceWith('powershell.exe', [
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', expect.any(String),
        ], {
            encoding: 'utf8', timeout: 15_000,
            env: { ...probe.fakeProcess.env, XLIDE_SOUNDS_FILE: file, XLIDE_SOUNDS_OWNER: '4321', XLIDE_SOUNDS_ACTION: 'restore' },
        });
        expect(probe.os.homedir).not.toHaveBeenCalled();
        const script = probe.spawnSync.mock.calls[0][1][4];
        expect(script).toContain('$file = $env:XLIDE_SOUNDS_FILE');
        expect(script).toContain('$owner = [int]$env:XLIDE_SOUNDS_OWNER');
        expect(script).toContain('$why = Invoke-OracleDialogSounds $env:XLIDE_SOUNDS_ACTION');
        // The worker's process runs both actions, so the function sets it.
        expect(script).toContain('$env:XLIDE_SOUNDS_ACTION = $Action');
        expect(script).not.toContain("O'Brien");
    });

    it.each(['throws', 'empty'])('falls back to os.homedir when userInfo %s', (mode) => {
        const probe = setup();
        if (mode === 'throws') { probe.os.userInfo.mockImplementation(() => { throw new Error('no account record'); }); }
        else { probe.os.userInfo.mockReturnValue({ homedir: '' }); }
        probe.setDialogSounds('restore');
        expect(probe.existsSync).toHaveBeenCalledExactlyOnceWith('C:\\Users\\fallback\\.xlide-vbe-oracle-sounds.json');
        expect(probe.spawnSync.mock.calls[0][2].env.XLIDE_SOUNDS_FILE).toBe('C:\\Users\\fallback\\.xlide-vbe-oracle-sounds.json');
        expect(probe.os.homedir).toHaveBeenCalled();
    });

    const failures: { name: string; completed: Completed; why: string }[] = [
        { name: 'nonzero', completed: { status: 1, stderr: ' access denied\r\n' }, why: 'access denied' },
        { name: 'timeout', completed: { status: null, error: new Error('spawnSync powershell.exe ETIMEDOUT'),
            stderr: 'ignored' }, why: 'spawnSync powershell.exe ETIMEDOUT' },
        { name: 'missing executable', completed: { status: null, error: new Error('spawnSync powershell.exe ENOENT') },
            why: 'spawnSync powershell.exe ENOENT' },
        { name: 'no diagnostic', completed: { status: 7, stderr: null }, why: 'PowerShell exited 7' },
        { name: 'held mutex', completed: { status: 3, stderr: 'another oracle run held the sounds for over ten seconds\r\n' },
            why: 'another oracle run held the sounds for over ten seconds' },
    ];
    const outcomes = [
        { outcome: 'accepted', expected: 'accepted', strict: false, code: 0 },
        { outcome: 'rejected', expected: 'rejected', strict: false, code: 0 },
        { outcome: 'rejected', expected: 'accepted', strict: false, code: 0 },
        { outcome: 'rejected', expected: 'accepted', strict: true, code: 1 },
        { outcome: 'worker_error', expected: 'accepted', strict: false, code: 1 },
    ];
    it.each(failures.flatMap((failure) => outcomes.map((result) => ({ ...failure, ...result }))))(
        'only warns for a restore that fails ($name); outcome=$outcome strict=$strict expected=$expected', async (row) => {
            const cases = [{ id: 'first', expected: row.expected }, { id: 'second', expected: row.expected }];
            const baseline = setup(cases, true, false);
            const probe = setup(cases);
            for (const current of [baseline, probe]) {
                current.runCaseWithRetries.mockImplementation(async (_worker, kase) => ({ caseId: kase.id, outcome: row.outcome }));
            }
            probe.spawnSync.mockReturnValue(row.completed);
            const argv = ['--json', ...(row.strict ? ['--strict'] : [])];
            expect(await probe.main(argv)).toBe(row.code);
            expect(await baseline.main(argv)).toBe(row.code);
            expect(probe.output).toEqual(baseline.output);
            expect(probe.runCaseWithRetries.mock.calls.map(([, kase]) => kase.id))
                .toEqual(baseline.runCaseWithRetries.mock.calls.map(([, kase]) => kase.id));
            expect(probe.errors).toEqual([`Could not put back the message-box sounds (${row.why}); the next oracle run will.`]);
            expect(probe.spawnSync).toHaveBeenCalledOnce();
            expectNoOwnedListeners(probe);
        });

    it('shows a notice the restore printed while succeeding', async () => {
        const probe = setup([{ id: 'first' }]);
        const notice = 'The saved message-box sounds could not be read; they were moved to C:\\x.unreadable-1.';
        probe.spawnSync.mockReturnValue({ status: 0, stderr: `${notice}\r\n` });
        expect(await probe.main(['--json'])).toBe(0);
        expect(probe.errors).toEqual([notice]);
    });
});

describe('VBE oracle worker sound reports', () => {
    const notice = 'The saved message-box sounds could not be read; they were moved to C:\\x.unreadable-1.';
    it.each([
        { name: 'a silence that failed', report: { action: 'silence', why: 'access denied', notice: '' },
            errors: ['Could not silence the message-box sounds for this run (access denied); rejected cases may play them.'] },
        { name: 'a silence with a notice', report: { action: 'silence', why: '', notice }, errors: [notice] },
        { name: 'a silence that failed with a notice', report: { action: 'silence', why: 'held', notice },
            errors: ['Could not silence the message-box sounds for this run (held); rejected cases may play them.'] },
        { name: 'a clean silence', report: { action: 'silence', why: '', notice: '' }, errors: [] },
        { name: 'a restore with a notice', report: { action: 'restore', why: '', notice }, errors: [notice] },
        // Left to the coordinator's restore, which finds the saved sounds still there.
        { name: 'a restore that failed', report: { action: 'restore', why: 'access denied', notice }, errors: [] },
        { name: 'another action', report: { action: 'other', why: 'x', notice }, errors: [] },
    ])('reports $name as the coordinator\'s own call did', ({ report, errors }) => {
        const probe = setup();
        probe.reportWorkerSounds(JSON.stringify(report));
        expect(probe.errors).toEqual(errors);
    });

    it.each(['', 'not json', 'null', '"text"'])('ignores an unreadable report %j', (text) => {
        const probe = setup();
        probe.reportWorkerSounds(text);
        expect(probe.errors).toEqual([]);
    });
});

// The real OracleWorker, with fake powershell.exe children that answer each
// request once they have reported the silence (sendSilence).
describe('VBE oracle worker start', () => {
    const start = source.indexOf('class OracleWorker {');
    const classSource = source.slice(start, source.indexOf('\n}', start) + 2);
    type Child = EventEmitter & { stdout: PassThrough; stderr: PassThrough; stdin: PassThrough; pid: number; requests: string[] };
    // silenceMs: how long a new worker may take to report its silence.
    function worker(silenceMs = 5000) {
        const children: Child[] = [];
        const spawn = vi.fn(() => {
            const child = Object.assign(new EventEmitter(), {
                stdout: new PassThrough(), stderr: new PassThrough(), stdin: new PassThrough(), pid: 70 + children.length, requests: [] as string[],
            });
            readline.createInterface({ input: child.stdin }).on('line', (line) => {
                child.requests.push(line);
                child.stdout.write(`XLIDE-ORACLE-RESULT {"caseId":${JSON.stringify(JSON.parse(line).casePath)}}\r\n`);
            });
            children.push(child);
            return child;
        });
        const spawnSync = vi.fn(() => ({ status: 0 }));
        const errors: string[] = [];
        const api = runInNewContext(`
            const RESULT_PREFIX = 'XLIDE-ORACLE-RESULT ';
            const WORKER = 'C:\\\\repo\\\\vbe_oracle_worker.ps1';
            const SOUNDS_SILENCE_MS = ${silenceMs};
            ${soundsConstant}
            ${prefixConstant}
            ${functions}
            ${workerSounds}
            class HostCleanupError extends Error {}
            ${classSource}
            ({ OracleWorker, workerCommand })`, {
            spawn, readline, spawnSync, setTimeout, clearTimeout,
            process: { pid: 4321, env: { PRESERVED: 'yes' } }, os: { userInfo: () => ({ homedir: 'C:\\Users\\account' }) },
            path: win32, console: { error: (line: string) => errors.push(line) },
        });
        const owned = { sweep: vi.fn(() => []), add: vi.fn() };
        const sendSilence = (child: Child, why = '') => child.stdout.write(`XLIDE-ORACLE-SOUNDS {"action":"silence","why":"${why}","notice":""}\r\n`);
        return { worker: new api.OracleWorker(7, owned), workerCommand: api.workerCommand, spawn, spawnSync, children, errors, sendSilence };
    }
    const tick = (ms = 0) => new Promise((resolve) => setTimeout(resolve, ms));

    it('runs the worker inside the command that silences and restores, with the sounds file and owner', async () => {
        const probe = worker();
        const reply = probe.worker.request({ casePath: 'c' }, 60_000);
        expect(probe.spawn).toHaveBeenCalledExactlyOnceWith('powershell.exe', [
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', probe.workerCommand(7),
        ], { stdio: ['pipe', 'pipe', 'pipe'], env: { PRESERVED: 'yes', XLIDE_SOUNDS_FILE: soundsFile,
            XLIDE_SOUNDS_OWNER: '4321', XLIDE_ORACLE_WORKER: 'C:\\repo\\vbe_oracle_worker.ps1', XLIDE_SOUNDS_SKIP: '' } });
        probe.sendSilence(probe.children[0]);
        expect(await reply).toEqual({ line: '{"caseId":"c"}', stderr: '' });
    });

    it('starts a case\'s time only once the worker has silenced the sounds, and reports how that went', async () => {
        const probe = worker();
        // A 150 ms case deadline, and a silence that takes 400 ms.
        const reply = probe.worker.request({ casePath: 'c' }, 150);
        await tick(400);
        expect(probe.children[0].requests).toEqual([]);
        probe.sendSilence(probe.children[0], 'access denied');
        expect(await reply).toEqual({ line: '{"caseId":"c"}', stderr: '' });
        expect(probe.errors).toEqual(['Could not silence the message-box sounds for this run (access denied); rejected cases may play them.']);
        // The next case goes to the same worker, with no wait.
        expect(await probe.worker.request({ casePath: 'd' }, 1000)).toEqual({ line: '{"caseId":"d"}', stderr: '' });
        expect(probe.spawn).toHaveBeenCalledOnce();
    });

    it('replaces a worker that does not report its silence in time with one that does not try', async () => {
        const probe = worker(200);
        const reply = probe.worker.request({ casePath: 'c' }, 1000);
        expect(await reply).toEqual({ line: '{"caseId":"c"}', stderr: '' });
        expect(probe.errors).toEqual(['Could not silence the message-box sounds for this run (no answer within 0.2 seconds); rejected cases may play them.']);
        expect(probe.spawnSync).toHaveBeenCalledExactlyOnceWith('taskkill', ['/PID', '70', '/T', '/F'], { encoding: 'utf8', timeout: 10_000 });
        expect(probe.spawn).toHaveBeenCalledTimes(2);
        expect((probe.spawn.mock.calls[1] as unknown[])[2]).toMatchObject({ env: { XLIDE_SOUNDS_SKIP: '1' } });
        expect(probe.children[0].requests).toEqual([]);
        expect(probe.children[1].requests).toEqual(['{"casePath":"c"}']);
        // Later workers of the run do not try either.
        probe.worker.kill();
        expect(await probe.worker.request({ casePath: 'd' }, 1000)).toEqual({ line: '{"caseId":"d"}', stderr: '' });
        expect((probe.spawn.mock.calls[2] as unknown[])[2]).toMatchObject({ env: { XLIDE_SOUNDS_SKIP: '1' } });
        expect(probe.errors).toHaveLength(1);
    });

    it('reports a worker that ends while silencing as one that ended during the case', async () => {
        const probe = worker();
        const reply = probe.worker.request({ casePath: 'c' }, 1000);
        probe.children[0].stderr.write('worker failed\r\n');
        await tick();
        probe.children[0].emit('close', 1);
        expect(await reply).toEqual({ exited: true, error: null, code: 1, stderr: 'worker failed' });
        expect(probe.children[0].requests).toEqual([]);
    });

    it('ignores the sound reports of a worker it has replaced', async () => {
        const probe = worker();
        const first = probe.worker.request({ casePath: 'c' }, 1000);
        probe.sendSilence(probe.children[0]);
        await first;
        probe.worker.kill();
        const second = probe.worker.request({ casePath: 'd' }, 1000);
        // The old worker's report neither shows nor counts as the new one's silence.
        probe.sendSilence(probe.children[0], 'late');
        await tick(20);
        expect(probe.children[1].requests).toEqual([]);
        probe.sendSilence(probe.children[1]);
        expect(await second).toEqual({ line: '{"caseId":"d"}', stderr: '' });
        expect(probe.errors).toEqual([]);
    });
});

// These assertions inspect the command argument AFTER JavaScript has evaluated
// String.raw and interpolated the events. They guard the emitted PS contract;
// they do not execute PowerShell or model Windows registry/JSON semantics.
// The restore's command and the worker's run the same function.
describe('VBE oracle generated dialog sound PowerShell', () => {
    function script() {
        const probe = setup();
        probe.setDialogSounds('restore');
        const restore: string = probe.spawnSync.mock.calls[0][1][4];
        const fn = restore.slice(restore.indexOf('function Invoke-OracleDialogSounds('), restore.indexOf('$why = Invoke-OracleDialogSounds $env:XLIDE_SOUNDS_ACTION'));
        expect(probe.workerCommand(40)).toContain(fn);
        return restore;
    }

    it('targets exactly the four fallback-related events and relative HKCU .Current keys', () => {
        const ps = script();
        expect(ps).toContain("$events = @('.Default', 'SystemAsterisk', 'SystemExclamation', 'SystemHand')");
        expect(ps).toContain(String.raw`$schemes = 'AppEvents\Schemes\Apps\.Default'`);
        expect(ps).toContain("[Microsoft.Win32.Registry]::CurrentUser.OpenSubKey((Join-Path (Join-Path $schemes $name) '.Current'), $writable)");
        expect(ps.match(/::CurrentUser\.OpenSubKey/g)).toHaveLength(1);
        expect(ps.match(/Open-Sound \$sound\.name \$true/g)).toHaveLength(2);
        expect(ps).toContain('$key = Open-Sound $name $false');
        expect(ps).not.toMatch(/HKEY_LOCAL_MACHINE|::LocalMachine|ExpandEnvironmentVariables/);
    });

    it('saves raw values and their kinds and uses the same kind for silence and restore', () => {
        const ps = script();
        expect(ps.startsWith("$ErrorActionPreference = 'Stop'\n")).toBe(true);
        expect(ps.match(/GetValue\('', \$null, 'DoNotExpandEnvironmentNames'\)/g)).toHaveLength(2);
        expect(ps).toContain("name = $name; value = $value; kind = $kind");
        expect(ps).toContain("$key.SetValue('', '', [Microsoft.Win32.RegistryValueKind]$sound.kind)");
        expect(ps).toContain("$key.SetValue('', [string]$sound.value, [Microsoft.Win32.RegistryValueKind]$sound.kind)");
        const restore = ps.slice(ps.indexOf('function Restore('), ps.indexOf('if (Test-Path'));
        // Only a value that still exists and is empty: one the user set or removed stays.
        expect(restore).toContain("if ($key.GetValueNames() -contains '' -and -not [string]$key.GetValue('', $null, 'DoNotExpandEnvironmentNames')) {");
        expect(restore).toContain('if ($null -eq $key) { continue }');
        const capture = ps.slice(ps.indexOf('$sounds = @()'), ps.indexOf('$json ='));
        expect(capture).toContain('if ($null -eq $key) { continue }');
        expect(capture).toContain("if ($value -and $kind -in @('String', 'ExpandString')) { $sounds += [pscustomobject]");
        expect(capture).toContain('if (-not $sounds.Count) { return }');
        expect(ps.match(/finally \{ \$key.Close\(\) \}/g)).toHaveLength(3);
    });

    it('keeps sounds nested in the saved object, sufficient JSON depth, and a singleton-safe restore loop', () => {
        const ps = script();
        expect(ps).toContain('$sounds = @()');
        expect(ps).toContain('foreach ($sound in @($saved.sounds))');
        // Windows PowerShell reads a file without a byte-order mark as the ANSI code page unless told.
        expect(ps).toContain('try { $saved = Get-Content -Raw -Encoding UTF8 -LiteralPath $file | ConvertFrom-Json } catch { }');
        expect(ps).toContain('$bytes = [Text.Encoding]::UTF8.GetBytes($json)');
        expect(ps).toContain('$snapshot = [pscustomobject]@{ owner = $owner; started = (Started $owner); sounds = $sounds }');
    });

    it('claims the saved-sounds file only if absent, before the first registry write', () => {
        const ps = script();
        const claim = ps.indexOf("try { $stream = [IO.File]::Open($file, 'CreateNew', 'Write', 'None') }");
        // Only "the file exists" means another run got there; any other failure is reported.
        expect(ps).toContain('if ($_.Exception.HResult -eq -2147024816) { return }\n        throw');
        const flush = ps.indexOf('$stream.Flush($true)');
        const silence = ps.lastIndexOf("$key.SetValue('', '',");
        expect(claim).toBeGreaterThan(0);
        expect(flush).toBeGreaterThan(claim);
        expect(silence).toBeGreaterThan(flush);
        expect(ps).not.toContain('.tmp');
    });

    it('does the whole save, repair or restore under one mutex per account', () => {
        const ps = script();
        // Every session of the account, elevated or not, opens the same mutex.
        expect(ps).toContain("$access.AddAccessRule((New-Object Security.AccessControl.MutexAccessRule($user, 'FullControl', 'Allow')))");
        const take = ps.indexOf('$mutex = [Threading.Mutex]::new($false, "Global\\xlide-vbe-oracle-sounds-$($user.Value)", [ref]$created, $access)');
        const wait = ps.indexOf('try { $owned = $mutex.WaitOne(10000) } catch [Threading.AbandonedMutexException] { $owned = $true }');
        // A wait that times out stops before touching the file or the registry.
        const timedOut = ps.indexOf("if (-not $owned) {\n        return 'another oracle run held the sounds for over ten seconds'\n    }");
        expect(timedOut).toBeGreaterThan(wait);
        // The restore's own process prints it and fails; the worker reports it.
        expect(ps).toContain("$why = Invoke-OracleDialogSounds $env:XLIDE_SOUNDS_ACTION\n    if ($why) {\n        [Console]::Error.WriteLine($why)\n        exit 3\n    }");
        expect(ps.indexOf('$file = $env:XLIDE_SOUNDS_FILE')).toBeGreaterThan(timedOut);
        // Messages reach Node as UTF-8, so a non-ASCII profile path reads correctly.
        expect(ps.indexOf('[Console]::OutputEncoding = [Text.Encoding]::UTF8')).toBeLessThan(take);
        const body = ps.indexOf('$file = $env:XLIDE_SOUNDS_FILE');
        const release = ps.indexOf('} finally { try { $mutex.ReleaseMutex() } catch { } }');
        expect(take).toBeGreaterThan(0);
        expect(wait).toBeGreaterThan(take);
        expect(body).toBeGreaterThan(wait);
        expect(release).toBeGreaterThan(ps.lastIndexOf("$key.SetValue('', '',"));
    });

    it('sets an unreadable saved-sounds file aside and goes on', () => {
        const ps = script();
        const unreadable = ps.indexOf('if (-not (Test-SavedSounds $saved)) {');
        // A file with no usable sound, or one entry it could not put back, is set aside too.
        expect(ps).toContain('$sound.value -isnot [string] -or [string]::IsNullOrEmpty($sound.value)');
        expect(unreadable).toBeGreaterThan(0);
        expect(ps).toContain("$aside = $file + '.unreadable-' + [DateTime]::UtcNow.Ticks");
        // A literal move: brackets in a profile path are not wildcards.
        expect(ps).toContain('[IO.File]::Move($file, $aside)');
        expect(ps).toContain('[Console]::Error.WriteLine("The saved message-box sounds could not be read; they were moved to $aside.")');
        expect(ps.indexOf('$sounds = @()')).toBeGreaterThan(unreadable);
    });

    it('checks pid AND process start time for restore and skips a live owner before stale recovery', () => {
        const ps = script();
        // An owner no longer listed has ended, whatever CIM can say.
        const listed = ps.indexOf("$listed = Get-Process -Id $id -ErrorAction SilentlyContinue\n        if ($null -eq $listed) { return '' }");
        expect(listed).toBeGreaterThan(0);
        expect(ps.indexOf('Get-CimInstance Win32_Process')).toBeGreaterThan(listed);
        expect(ps).toContain('try { $time = $listed.StartTime.ToFileTimeUtc() }');
        // An elevated owner a normal shell cannot open: its creation time through CIM.
        expect(ps).toContain('try { $process = Get-CimInstance Win32_Process -Filter "ProcessId=$id" -ErrorAction Stop } catch { return \'unreadable\' }');
        expect(ps).toContain("if ($null -eq $process) { return '' }");
        expect(ps).toContain("try { $time = $process.CreationDate.ToFileTimeUtc() } catch { return 'unreadable' }");
        // CIM has microseconds only, so both are compared to the microsecond.
        expect(ps).toContain('return [string]($time - ($time % 10))');
        // A restore puts back only this run's own saved sounds...
        const own = ps.indexOf("} elseif ($env:XLIDE_SOUNDS_ACTION -eq 'restore' -and $saved.owner -eq $owner -and $saved.started -eq (Started $owner)) {\n            Restore $saved\n            Remove-Item -Force -LiteralPath $file\n            return\n");
        // ...and a running owner, or one that cannot be checked, is left alone.
        const skip = ps.indexOf("} elseif ($saved.started -and (Started $saved.owner) -in @($saved.started, 'unreadable')) {");
        const recover = ps.indexOf('# Left by a run that was killed');
        // A file that cannot be removed stops the script, so Node warns, rather
        // than leaving it to look like another run's.
        expect(ps).toContain('# so a run that starts no worker still puts them back.\n            Restore $saved\n            Remove-Item -Force -LiteralPath $file\n');
        expect(own).toBeGreaterThan(0);
        expect(skip).toBeGreaterThan(own);
        expect(recover).toBeGreaterThan(skip);
        expect(ps.indexOf("if ($env:XLIDE_SOUNDS_ACTION -eq 'restore') { return }")).toBeGreaterThan(recover);
        expect(ps.indexOf('$sounds = @()')).toBeGreaterThan(recover);
    });
});

describe('VBE oracle generated worker command', () => {
    const ps = setup().workerCommand(40);

    it('silences before the worker, and puts back in finally, in one process', () => {
        expect(ps.startsWith("$ErrorActionPreference = 'Stop'\n")).toBe(true);
        const skip = ps.indexOf('$silence = -not $env:XLIDE_SOUNDS_SKIP\n');
        const silence = ps.indexOf("if ($silence) { Send-OracleDialogSounds 'silence' }\n");
        const worker = ps.indexOf('try { & $env:XLIDE_ORACLE_WORKER -Serve -CasesPerInstance 40 }\n');
        const restore = ps.indexOf("finally { if ($silence) { Send-OracleDialogSounds 'restore' } }");
        expect(skip).toBeGreaterThan(ps.indexOf('function Send-OracleDialogSounds('));
        expect(silence).toBeGreaterThan(skip);
        expect(worker).toBeGreaterThan(silence);
        expect(restore).toBeGreaterThan(worker);
        expect(ps.trimEnd().endsWith("finally { if ($silence) { Send-OracleDialogSounds 'restore' } }")).toBe(true);
        // The worker keeps its own output encoding: nothing here changes it.
        expect(ps).not.toContain('OutputEncoding');
        // An exit would end the process before the worker or its finally.
        expect(ps).not.toMatch(/\bexit\b/);
    });

    it('sends each report as one ASCII line and keeps notices off the worker stderr', () => {
        const send = ps.slice(ps.indexOf('function Send-OracleDialogSounds('), ps.indexOf('$silence = -not $env:XLIDE_SOUNDS_SKIP'));
        expect(send).toContain('[Console]::SetError($notices)');
        expect(send).toContain('finally { [Console]::SetError($stderr) }');
        // A failure the function throws is a reason too; it never stops the worker.
        expect(send).toContain('try { $why = Invoke-OracleDialogSounds $Action }\n        catch { $why = ($_ | Out-String).Trim() }');
        expect(send).toContain("if (-not $why) { $why = '' }");
        expect(send).toContain("[regex]::Replace($report, '[^\\x00-\\x7F]', { param($m) '\\u{0:x4}' -f [int][char]$m.Value })");
        expect(send).toContain("[Console]::Out.WriteLine('XLIDE-ORACLE-SOUNDS ' + $report)\n        [Console]::Out.Flush()");
        expect(source).toContain("const SOUNDS_PREFIX = 'XLIDE-ORACLE-SOUNDS ';");
    });
});

// The worker command, executed on Windows with a stub worker and the registry
// and process lookups replaced by in-memory fakes, as in
// vbeOracleSoundRecovery.test.ts. One PowerShell start per scenario.
describe.skipIf(process.platform !== 'win32')('executed worker command', () => {
    function run(mode: 'serve' | 'throw' | 'skip') {
        const root = mkdtempSync(join(tmpdir(), 'xlide-worker-sounds-test-'));
        try {
            let ps = setup().workerCommand(9);
            const openBegin = ps.indexOf('    function Open-Sound(');
            const openEnd = ps.indexOf('    # A normal shell', openBegin);
            ps = ps.slice(0, openBegin) + String.raw`
    function Open-Sound($name, $writable) {
        $key = [pscustomobject]@{ Name = [string]$name }
        $key | Add-Member ScriptMethod GetValueNames { return ,@('') }
        $key | Add-Member ScriptMethod GetValue { param($n, $d, $o); return $global:fakeSounds[$this.Name] }
        $key | Add-Member ScriptMethod GetValueKind { param($n); return 'String' }
        $key | Add-Member ScriptMethod SetValue { param($n, $value, $kind); $global:fakeSounds[$this.Name] = $value; [Console]::Out.WriteLine("XLIDE-TEST-WRITE $($this.Name)=$value") }
        $key | Add-Member ScriptMethod Close { }
        return $key
    }
` + ps.slice(openEnd);
            const startedBegin = ps.indexOf('    function Started(');
            const startedEnd = ps.indexOf('    # Only a sound', startedBegin);
            ps = ps.slice(0, startedBegin) + "    function Started($id) { return '1234567890' }\n" + ps.slice(startedEnd);
            // A mutex of the test's own, not the account's.
            const mutex = `"Global\\xlide-test-${process.pid}-${mode}"`;
            ps = ps.replace('"Global\\xlide-vbe-oracle-sounds-$($user.Value)"', mutex);
            expect(ps).toContain(mutex);
            expect(ps).not.toMatch(/::CurrentUser|Get-CimInstance|Get-Process/);
            ps = `$global:fakeSounds = @{ '.Default' = 'ding.wav'; 'SystemAsterisk' = 'chord.wav'; 'SystemExclamation' = ''; 'SystemHand' = 'stop.wav' }\n${ps}`;
            const stub = join(root, 'worker.ps1');
            writeFileSync(stub, String.raw`param([switch]$Serve, [int]$CasesPerInstance)
$ErrorActionPreference = 'Stop'
[Console]::Out.WriteLine("XLIDE-TEST-DURING " + (($global:fakeSounds.Keys | Sort-Object | ForEach-Object { "$_=$($global:fakeSounds[$_])" }) -join ','))
while ($null -ne ($line = [Console]::In.ReadLine())) {
    if ($line -eq 'throw') { throw 'stub worker failed' }
    [Console]::Out.WriteLine("XLIDE-ORACLE-RESULT $line serve=$Serve per=$CasesPerInstance")
}
`);
            const file = join(root, 'sounds.json');
            const completed = realSpawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', ps], {
                encoding: 'utf8', timeout: 30_000, input: mode === 'throw' ? 'one\nthrow\n' : 'one\ntwo\n',
                env: { ...process.env, XLIDE_SOUNDS_FILE: file, XLIDE_SOUNDS_OWNER: '4321', XLIDE_ORACLE_WORKER: stub,
                    XLIDE_SOUNDS_SKIP: mode === 'skip' ? '1' : '' },
            });
            return { ...completed, lines: completed.stdout.split(/\r?\n/).filter(Boolean), left: existsSync(file) };
        } finally { rmSync(root, { recursive: true, force: true }); }
    }
    const silenced = ['XLIDE-TEST-WRITE .Default=', 'XLIDE-TEST-WRITE SystemAsterisk=', 'XLIDE-TEST-WRITE SystemHand='];
    const report = (action: string) => `XLIDE-ORACLE-SOUNDS {"action":"${action}","why":"","notice":""}`;
    const during = 'XLIDE-TEST-DURING .Default=,SystemAsterisk=,SystemExclamation=,SystemHand=';
    const restored = ['XLIDE-TEST-WRITE .Default=ding.wav', 'XLIDE-TEST-WRITE SystemAsterisk=chord.wav', 'XLIDE-TEST-WRITE SystemHand=stop.wav'];

    it('silences before the worker serves, and puts back after it ends', () => {
        const result = run('serve');
        expect(result.error).toBeUndefined();
        expect(result.status, result.stderr).toBe(0);
        expect(result.stderr).toBe('');
        expect(result.lines).toEqual([...silenced, report('silence'), during,
            'XLIDE-ORACLE-RESULT one serve=True per=9', 'XLIDE-ORACLE-RESULT two serve=True per=9', ...restored, report('restore')]);
        expect(result.left).toBe(false);
    }, 40_000);

    it('puts back after a worker that fails, and keeps its error and exit code', () => {
        const result = run('throw');
        expect(result.status).toBe(1);
        expect(result.stderr).toContain('stub worker failed');
        expect(result.lines).toEqual([...silenced, report('silence'), during,
            'XLIDE-ORACLE-RESULT one serve=True per=9', ...restored, report('restore')]);
        expect(result.left).toBe(false);
    }, 40_000);

    it('leaves the sounds alone in a worker started after one did not report in time', () => {
        const result = run('skip');
        expect(result.status, result.stderr).toBe(0);
        expect(result.lines).toEqual(['XLIDE-TEST-DURING .Default=ding.wav,SystemAsterisk=chord.wav,SystemExclamation=,SystemHand=stop.wav',
            'XLIDE-ORACLE-RESULT one serve=True per=9', 'XLIDE-ORACLE-RESULT two serve=True per=9']);
        expect(result.left).toBe(false);
    }, 40_000);
});

// How many times a run starts powershell.exe, with the real host probe, worker
// class and restore, and fake processes. The owner asked for as few as
// possible: one probe for every host, and one worker that also silences and
// restores the sounds.
describe('VBE oracle PowerShell starts per run', () => {
    // The real host probes, combined probe included.
    const probeSource = ['let hostAnswers;', ...['probedHost', 'probeHosts', 'comRegistered', 'registeredHostBuild'].map((name) => {
        const start = source.indexOf(`function ${name}(`);
        return source.slice(start, source.indexOf('\n}', start) + 2);
    })].join('\n');
    const classStart = source.indexOf('class OracleWorker {');
    const classSource = source.slice(classStart, source.indexOf('\n}', classStart) + 2);
    const build = '16.0.20527.20196';

    async function count(cases: Case[], { soundsLeft = false, checkpoint = false, saved = [] as string[] } = {}) {
        const starts: string[] = [];
        const spawnSync = vi.fn((command: string, args: string[]) => {
            if (command !== 'powershell.exe') { return { status: 0 }; }
            const script = args[4];
            if (!script.includes('XLIDE-ORACLE-HOST')) { starts.push('restore'); return { status: 0, stderr: '' }; }
            starts.push('probe');
            const progIds = /foreach \(\$progId in @\(([^)]*)\)\)/.exec(script)![1].split(', ').map((progId) => progId.slice(1, -1));
            return { status: 0, stdout: progIds.map((progId) => `XLIDE-ORACLE-HOST ${progId} registered\r\nXLIDE-ORACLE-HOST ${progId} build ${build}\r\n`).join('') };
        });
        const spawn = vi.fn((command: string) => {
            starts.push(command === 'powershell.exe' ? 'worker' : command);
            const child = Object.assign(new EventEmitter(), {
                stdout: new PassThrough(), stderr: new PassThrough(), stdin: new PassThrough(), pid: 77,
            });
            readline.createInterface({ input: child.stdin }).on('line', (line) => {
                const request = JSON.parse(line);
                child.stdout.write(`XLIDE-ORACLE-RESULT ${JSON.stringify({ caseId: request.caseId, outcome: 'accepted', hostBuild: build })}\r\n`);
            });
            child.stdin.on('finish', () => setImmediate(() => child.emit('close', 0)));
            child.stdout.write('XLIDE-ORACLE-SOUNDS {"action":"silence","why":"","notice":""}\r\n');
            return child;
        });
        const fakeProcess = Object.assign(new EventEmitter(), { platform: 'win32', pid: 4321, env: {}, exit: vi.fn() });
        const main = runInNewContext(`
            const DEFAULT_HOST = 'excel';
            const DEFAULT_CASES = 'cases.json';
            const DEFAULT_TIMEOUT_RETRIES = 2;
            const DEFAULT_CASES_PER_INSTANCE = 40;
            const USAGE = 'usage';
            const RESULT_PREFIX = 'XLIDE-ORACLE-RESULT ';
            const WORKER = 'worker.ps1';
            const SOUNDS_SILENCE_MS = 15_000;
            const EVIDENCE_OUTCOMES = new Set(['accepted', 'rejected']);
            ${soundsConstant}
            ${prefixConstant}
            ${functions}
            ${workerSounds}
            ${probeSource}
            class HostCleanupError extends Error {}
            ${classSource}
            ${mainSource}
            main`, {
            process: fakeProcess, os: { userInfo: () => ({ homedir: 'C:\\Users\\account' }) }, path: win32,
            spawn, spawnSync, readline, setTimeout, clearTimeout,
            console: { log: () => {}, error: () => {} },
            OFFICE_HOSTS: Object.fromEntries(Object.entries(OFFICE_HOST_APPS).map(([host, info]) => [host,
                { progId: info.progId, image: info.processName, noun: info.noun }])),
            loadCasesDocument: () => ({ cases }),
            // A checkpoint that holds the listed cases.
            loadCheckpoint: () => new Map(saved.map((id) => [id, { caseId: id, outcome: 'accepted' }])),
            caseHash: () => 'hash', CHECKPOINT_KEY_VERSION: 2, checkpointKey: (id: string) => id, phaseMatches: () => true,
            acquireOracleRunLock: async () => ({ held: true, release: async () => {} }),
            oracleRunLockAddress: () => 'test-lock', oracleRunOwner: () => ({ pid: 1 }), describeRefusal: () => 'refused',
            OwnedHosts: class { load() { return []; } add() {} sweep() { return []; } },
            ownedHostsPath: () => 'hosts.json', describeLeftHosts: () => 'left', abandonActiveCases: () => {},
            unwritableCheckpoint: () => undefined,
            runCaseWithRetries: async (worker: { request: (m: object, t: number) => Promise<{ line: string }> }, kase: Case) =>
                JSON.parse((await worker.request({ caseId: kase.id }, 60_000)).line),
            fs: { appendFileSync: vi.fn(), existsSync: () => soundsLeft },
        });
        expect(await main(['--json', ...(checkpoint ? ['--checkpoint', 'checkpoint.jsonl'] : [])])).toBe(0);
        return starts;
    }
    const everyHost = Object.keys(OFFICE_HOST_APPS).flatMap((host) => [{ id: `${host}-1`, host }, { id: `${host}-2`, host }]);

    it('starts two for a run of every host: one probe and one worker', async () => {
        expect(await count(everyHost)).toEqual(['probe', 'worker']);
    });

    it('starts two for a checkpoint run of every host, the builds coming from the probe', async () => {
        expect(await count(everyHost, { checkpoint: true })).toEqual(['probe', 'worker']);
    });

    it('starts two for a single case', async () => {
        expect(await count([{ id: 'only', host: 'word' }])).toEqual(['probe', 'worker']);
    });

    it('starts a restore only when the worker left the saved sounds', async () => {
        expect(await count(everyHost, { soundsLeft: true })).toEqual(['probe', 'worker', 'restore']);
    });

    it('starts only the probe for a run whose cases all come from the checkpoint', async () => {
        expect(await count(everyHost, { checkpoint: true, saved: everyHost.map((kase) => kase.id) })).toEqual(['probe']);
    });
});
