import { EventEmitter } from 'node:events';
import { readFileSync } from 'node:fs';
import { join, win32 } from 'node:path';
import { runInNewContext } from 'node:vm';
import { describe, expect, it, vi } from 'vitest';
import { OFFICE_HOST_APPS } from '../src/officeHostApps';

// Run the coordinator's actual functions, with no real PowerShell, signals,
// registry writes, Office processes or account-profile writes. The run lock
// and the host clean-up have tests of their own (vbeOracleLock.test.ts).
const source = readFileSync(join(process.cwd(), 'syntax_corpus/oracle/run_vbe_oracle.mjs'), 'utf8');
const names = ['dialogSoundsPath', 'dialogSoundsScript', 'setDialogSounds', 'parseArgs',
    'caseHost', 'hostsNeedingBuilds', 'knownHostCrashResult', 'expectedMatches', 'isOracleInfrastructureFailure'];
const functions = names.map((name) => {
    const start = source.indexOf(`function ${name}(`);
    // Up to the next top-level declaration, so a closing brace at the start of
    // a line inside a template literal cannot cut a function short.
    const next = source.slice(start + 1).search(/\n(?:async function |function |const |let |\/\/ )/);
    return next < 0 ? source.slice(start) : source.slice(start, start + 1 + next);
}).join('\n');
const soundsConstant = source.match(/^const DIALOG_SOUNDS = .*;$/m)![0];
const mainStart = source.indexOf('async function main(');
const mainSource = source.slice(mainStart, source.indexOf('\nprocess.exit(', mainStart));
const signals = [['SIGINT', 130], ['SIGBREAK', 149], ['SIGHUP', 129]] as const;
const ownedEvents = ['exit', ...signals.map(([signal]) => signal)];
type Case = { id: string; host?: string; expected?: string };
type Completed = { status: number | null; stderr?: string | null; error?: Error };
type SpawnOptions = { encoding: string; timeout: number; env: Record<string, string> };

function deferred<T>() {
    let resolve!: (value: T) => void;
    let reject!: (reason: Error) => void;
    const promise = new Promise<T>((yes, no) => { resolve = yes; reject = no; });
    return { promise, resolve, reject };
}

function setup(cases: Case[] = [{ id: 'first' }, { id: 'second' }]) {
    const trace: string[] = [];
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
    const spawnSync = vi.fn((_command: string, _args: string[], options: SpawnOptions): Completed => {
        trace.push(options.env.XLIDE_SOUNDS_ACTION);
        return { status: 0 };
    });
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
        ${functions}
        ${mainSource}
        ({ main, dialogSoundsPath, setDialogSounds })`, {
        process: fakeProcess, os, path: win32, spawnSync,
        console: { log: (line: string) => output.push(line), error: (line: string) => errors.push(line) },
        OFFICE_HOSTS: Object.fromEntries(Object.entries(OFFICE_HOST_APPS).map(([host, info]) => [host,
            { progId: info.progId, image: info.processName, noun: info.noun }])),
        loadCasesDocument, comRegistered,
        registeredHostBuild,
        acquireOracleLock: () => () => {}, oracleLockPath: () => 'C:\\Users\\account\\.xlide-vbe-oracle.lock',
        killProcess: () => true,
        loadCheckpoint: () => new Map(),
        OracleWorker: class {
            async close() {
                trace.push('close:begin');
                try { await close(); } finally { trace.push('close:end'); }
            }
            kill() { kill(); }
        },
        caseHash: () => 'hash', checkpointKey: (id: string) => id, phaseMatches: () => true,
        runCaseWithRetries: (worker: unknown, kase: Case) => {
            trace.push(`measure:${kase.id}`);
            return runCaseWithRetries(worker, kase);
        },
        fs: { appendFileSync: vi.fn() },
    }) as {
        main: (argv: string[]) => Promise<number>;
        dialogSoundsPath: () => string;
        setDialogSounds: (action: string) => void;
    };
    return { ...api, trace, errors, output, fakeProcess, os, spawnSync, close, kill,
        runCaseWithRetries, comRegistered, registeredHostBuild, loadCasesDocument, interrupted };
}

function expectNoOwnedListeners(probe: ReturnType<typeof setup>) {
    for (const event of ownedEvents) { expect(probe.fakeProcess.listenerCount(event), event).toBe(0); }
}

describe('VBE oracle dialog sound lifecycle', () => {
    it.each(Object.keys(OFFICE_HOST_APPS))('silences once before measuring %s, restores once after close, and removes its listeners', async (host) => {
        const probe = setup([{ id: 'first', host }, { id: 'second', host }]);
        const unrelated = vi.fn();
        for (const event of ownedEvents) { probe.fakeProcess.on(event, unrelated); }
        expect(await probe.main(['--json'])).toBe(0);
        expect(probe.trace).toEqual(['silence', 'measure:first', 'measure:second', 'close:begin', 'close:end', 'restore']);
        expect(probe.errors).toEqual([]);
        for (const event of ownedEvents) { expect(probe.fakeProcess.listeners(event)).toEqual([unrelated]); }
        probe.fakeProcess.emit('exit', 0);
        expect(probe.spawnSync).toHaveBeenCalledTimes(2);
    });

    it.each(['case', 'close', 'both'])('restores and removes listeners when %s throws', async (failure) => {
        const probe = setup();
        const caseError = new Error('case failed');
        const closeError = new Error('close failed');
        if (failure !== 'close') { probe.runCaseWithRetries.mockRejectedValueOnce(caseError); }
        if (failure !== 'case') { probe.close.mockRejectedValueOnce(closeError); }
        await expect(probe.main([])).rejects.toBe(failure === 'case' ? caseError : closeError);
        expect(probe.trace).toEqual(['silence', 'measure:first',
            ...(failure === 'close' ? ['measure:second'] : []), 'close:begin', 'close:end', 'restore']);
        expect(probe.close).toHaveBeenCalledOnce();
        expectNoOwnedListeners(probe);
        probe.fakeProcess.emit('exit', 1);
        expect(probe.spawnSync).toHaveBeenCalledTimes(2);
    });

    it('restores synchronously on exit, without restoring again in finally', async () => {
        const probe = setup([{ id: 'first' }]);
        const pending = deferred<{ caseId: string; outcome: string }>();
        probe.runCaseWithRetries.mockReturnValueOnce(pending.promise);
        const running = probe.main(['--json']);
        expect(probe.trace).toEqual(['silence', 'measure:first']);
        probe.fakeProcess.emit('exit', 1);
        // The clean-up ends the worker, then puts the sounds back.
        expect(probe.trace).toEqual(['silence', 'measure:first', 'kill', 'restore']);
        probe.fakeProcess.emit('exit', 1);
        pending.resolve({ caseId: 'first', outcome: 'accepted' });
        expect(await running).toBe(0);
        expect(probe.trace).toEqual(['silence', 'measure:first', 'kill', 'restore', 'close:begin', 'close:end']);
        expectNoOwnedListeners(probe);
    });

    it.each(signals.flatMap(([signal, code]) => ['case', 'close'].map((stage) => ({ signal, code, stage }))))(
        'restores and exits $code on $signal during $stage, after ending the worker, without a second restore', async ({ signal, code, stage }) => {
            const probe = setup([{ id: 'first' }]);
            const measuring = deferred<{ caseId: string; outcome: string }>();
            const closing = deferred<void>();
            const closeStarted = deferred<void>();
            probe.runCaseWithRetries.mockReturnValueOnce(measuring.promise);
            probe.close.mockImplementationOnce(() => { closeStarted.resolve(); return closing.promise; });
            const running = probe.main(['--json']);
            const rejected = expect(running).rejects.toBe(probe.interrupted);
            if (stage === 'close') {
                measuring.resolve({ caseId: 'first', outcome: 'accepted' });
                await closeStarted.promise;
            }
            expect(() => probe.fakeProcess.emit(signal)).toThrow(probe.interrupted);
            expect(probe.trace.slice(-3)).toEqual(['kill', 'restore', `exit:${code}`]);
            expect(probe.fakeProcess.exit).toHaveBeenCalledExactlyOnceWith(code);
            expect(probe.kill).toHaveBeenCalledOnce();
            // Still listening until the run ends, so a second one cannot end the run mid-clean-up.
            expect(probe.fakeProcess.listenerCount(signal)).toBe(1);
            if (stage === 'case') {
                measuring.reject(probe.interrupted);
                closing.resolve();
            } else { closing.reject(probe.interrupted); }
            await rejected;
            expect(probe.trace.filter((event) => event === 'restore')).toHaveLength(1);
            expect(probe.spawnSync).toHaveBeenCalledTimes(2);
            expectNoOwnedListeners(probe);
        });

    it.each(signals)('restores once and exits %s on %s arriving while the sounds are being silenced', async (signal, code) => {
        const probe = setup([{ id: 'first' }]);
        // The signal is handled from inside the silencing call, as a Ctrl-C
        // reaching the run while that PowerShell runs would be once it returns.
        probe.spawnSync.mockImplementationOnce((_command, _args, options) => {
            probe.trace.push(options.env.XLIDE_SOUNDS_ACTION);
            probe.fakeProcess.emit(signal);
            return { status: 0 };
        });
        await expect(probe.main(['--json'])).rejects.toBe(probe.interrupted);
        expect(probe.trace).toEqual(['silence', 'kill', 'restore', `exit:${code}`, 'close:begin', 'close:end']);
        expect(probe.runCaseWithRetries).not.toHaveBeenCalled();
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
        { argv: [], platform: 'linux', code: 2 },
        { argv: [], platform: 'darwin', code: 2 },
        { argv: ['--help'], platform: 'win32', code: 0 },
    ])('does not touch sounds for $argv on $platform', async ({ argv, platform, code }) => {
        const probe = setup();
        probe.fakeProcess.platform = platform;
        expect(await probe.main(argv)).toBe(code);
        expect(probe.spawnSync).not.toHaveBeenCalled();
        expect(probe.runCaseWithRetries).not.toHaveBeenCalled();
        expectNoOwnedListeners(probe);
    });

    it.each(['unknown host', 'unregistered host', 'invalid document', 'unreadable checkpoint build'])(
        'does not touch sounds after startup failure: %s', async (failure) => {
            const probe = setup([{ id: 'first', host: failure === 'unknown host' ? 'other' : 'excel' }]);
            if (failure === 'unregistered host') { probe.comRegistered.mockReturnValue(false); }
            if (failure === 'invalid document') { probe.loadCasesDocument.mockImplementation(() => { throw new Error('bad JSON'); }); }
            if (failure === 'unreadable checkpoint build') {
                probe.registeredHostBuild.mockReturnValue({ registered: true, build: null, why: 'unreadable' });
            }
            expect(await probe.main(failure === 'unreadable checkpoint build' ? ['--checkpoint', 'checkpoint'] : [])).toBe(2);
            expect(probe.spawnSync).not.toHaveBeenCalled();
            expectNoOwnedListeners(probe);
        });
});

describe('VBE oracle dialog sound PowerShell invocation', () => {
    it.each(['silence', 'restore'])('passes account-profile path, pid, action and inherited environment for %s', (action) => {
        const probe = setup();
        probe.os.userInfo.mockReturnValue({ homedir: "C:\\Users\\O'Brien [account]" });
        probe.setDialogSounds(action);
        expect(probe.spawnSync).toHaveBeenCalledExactlyOnceWith('powershell.exe', [
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', expect.any(String),
        ], {
            encoding: 'utf8', timeout: 15_000,
            env: { ...probe.fakeProcess.env,
                XLIDE_SOUNDS_FILE: "C:\\Users\\O'Brien [account]\\.xlide-vbe-oracle-sounds.json",
                XLIDE_SOUNDS_OWNER: '4321', XLIDE_SOUNDS_ACTION: action },
        });
        expect(probe.os.homedir).not.toHaveBeenCalled();
        const script = probe.spawnSync.mock.calls[0][1][4];
        expect(script).toContain('$file = $env:XLIDE_SOUNDS_FILE');
        expect(script).toContain('$owner = [int]$env:XLIDE_SOUNDS_OWNER');
        expect(script).not.toContain("O'Brien");
    });

    it.each(['throws', 'empty'])('falls back to os.homedir when userInfo %s', (mode) => {
        const probe = setup();
        if (mode === 'throws') { probe.os.userInfo.mockImplementation(() => { throw new Error('no account record'); }); }
        else { probe.os.userInfo.mockReturnValue({ homedir: '' }); }
        probe.setDialogSounds('silence');
        expect(probe.spawnSync.mock.calls[0][2].env.XLIDE_SOUNDS_FILE).toBe('C:\\Users\\fallback\\.xlide-vbe-oracle-sounds.json');
        expect(probe.os.homedir).toHaveBeenCalledOnce();
    });

    const failures: { name: string; completed: Completed; why: string }[] = [
        { name: 'nonzero', completed: { status: 1, stderr: ' access denied\r\n' }, why: 'access denied' },
        { name: 'timeout', completed: { status: null, error: new Error('spawnSync powershell.exe ETIMEDOUT'),
            stderr: 'ignored' }, why: 'spawnSync powershell.exe ETIMEDOUT' },
        { name: 'missing executable', completed: { status: null, error: new Error('spawnSync powershell.exe ENOENT') },
            why: 'spawnSync powershell.exe ENOENT' },
        { name: 'no diagnostic', completed: { status: 7, stderr: null }, why: 'PowerShell exited 7' },
    ];
    const outcomes = [
        { outcome: 'accepted', expected: 'accepted', strict: false, code: 0 },
        { outcome: 'rejected', expected: 'rejected', strict: false, code: 0 },
        { outcome: 'rejected', expected: 'accepted', strict: false, code: 0 },
        { outcome: 'rejected', expected: 'accepted', strict: true, code: 1 },
        { outcome: 'worker_error', expected: 'accepted', strict: false, code: 1 },
    ];
    const failureScenarios = failures.flatMap((failure) => ['silence', 'restore', 'both'].flatMap((action) => outcomes.map((result) => ({
        ...failure, action, ...result,
    }))));
    it.each(failureScenarios)('only warns for $name during $action; outcome=$outcome strict=$strict expected=$expected', async (row) => {
        const cases = [{ id: 'first', expected: row.expected }, { id: 'second', expected: row.expected }];
        const baseline = setup(cases);
        const probe = setup(cases);
        for (const current of [baseline, probe]) {
            current.runCaseWithRetries.mockImplementation(async (_worker, kase) => ({ caseId: kase.id, outcome: row.outcome }));
        }
        probe.spawnSync.mockImplementation((_command, _args, options) => {
            const action = options.env.XLIDE_SOUNDS_ACTION;
            probe.trace.push(action);
            return row.action === 'both' || row.action === action ? row.completed : { status: 0 };
        });
        const argv = ['--json', ...(row.strict ? ['--strict'] : [])];
        expect(await probe.main(argv)).toBe(row.code);
        expect(await baseline.main(argv)).toBe(row.code);
        expect(probe.output).toEqual(baseline.output);
        expect(probe.runCaseWithRetries.mock.calls.map(([, kase]) => kase.id))
            .toEqual(baseline.runCaseWithRetries.mock.calls.map(([, kase]) => kase.id));
        expect(probe.errors).toEqual([
            ...(row.action !== 'restore' ? [`Could not silence the message-box sounds for this run (${row.why}); rejected cases may play them.`] : []),
            ...(row.action !== 'silence' ? [`Could not put back the message-box sounds (${row.why}); the next oracle run will.`] : []),
        ]);
        expect(probe.spawnSync).toHaveBeenCalledTimes(2);
        expectNoOwnedListeners(probe);
    });

    it('shows a notice the sound script printed while succeeding', async () => {
        const probe = setup([{ id: 'first' }]);
        probe.runCaseWithRetries.mockImplementation(async (_worker, kase) => ({ caseId: kase.id, outcome: 'accepted' }));
        const notice = 'The saved message-box sounds could not be read; they were moved to C:\\x.unreadable-1.';
        probe.spawnSync.mockImplementation((_command, _args, options) => (options.env.XLIDE_SOUNDS_ACTION === 'silence'
            ? { status: 0, stderr: `${notice}\r\n` } : { status: 0, stderr: '' }));
        expect(await probe.main(['--json'])).toBe(0);
        expect(probe.errors).toEqual([notice]);
    });
});

// These assertions inspect the command argument AFTER JavaScript has evaluated
// String.raw and interpolated the events. They guard the emitted PS contract;
// they do not execute PowerShell or model Windows registry/JSON semantics.
describe('VBE oracle generated dialog sound PowerShell', () => {
    function script() {
        const probe = setup();
        probe.setDialogSounds('silence');
        return probe.spawnSync.mock.calls[0][1][4];
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
        expect(capture).toContain('if (-not $sounds.Count) { exit 0 }');
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
        expect(ps).toContain('if ($_.Exception.HResult -eq -2147024816) { exit 0 }\n        throw');
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
        const timedOut = ps.indexOf("if (-not $owned) {\n        [Console]::Error.WriteLine('another oracle run held the sounds for over ten seconds')\n        exit 3\n    }");
        expect(timedOut).toBeGreaterThan(wait);
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
        expect(ps).toContain('if ($saved.owner -eq $owner -and $saved.started -eq (Started $owner)) {');
        // A running owner that cannot be checked is left alone too.
        const skip = ps.indexOf("} elseif ($saved.started -and (Started $saved.owner) -in @($saved.started, 'unreadable')) {");
        const recover = ps.indexOf('# Left by a run that was killed.');
        // A file that cannot be removed stops the script, so Node warns, rather
        // than leaving it to look like another run's.
        expect(ps).toContain('# Left by a run that was killed.\n            Restore $saved\n            Remove-Item -Force -LiteralPath $file\n');
        expect(skip).toBeGreaterThan(0);
        expect(recover).toBeGreaterThan(skip);
        expect(ps.indexOf("if ($env:XLIDE_SOUNDS_ACTION -eq 'restore') { exit 0 }")).toBeGreaterThan(recover);
        expect(ps.indexOf('$sounds = @()')).toBeGreaterThan(recover);
    });
});
