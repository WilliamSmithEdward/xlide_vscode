import * as fs from 'node:fs';
import * as path from 'node:path';
import { EventEmitter } from 'node:events';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { OFFICE_HOST_APPS } from '../src/officeHostApps';
import { loadCoordinator, oracleDir, removeTemporary } from './helpers/vbeOracleCoordinator';

afterEach(() => {
    vi.useRealTimers();
    removeTemporary();
});

// The real worker class, with PowerShell's process and output replaced.
function setup() {
    const lines = new EventEmitter();
    const child = Object.assign(new EventEmitter(), {
        pid: 4343,
        stdout: { setEncoding() {} },
        stderr: Object.assign(new EventEmitter(), { setEncoding() {} }),
        stdin: Object.assign(new EventEmitter(), { write: vi.fn(), end: vi.fn() }),
        kill: vi.fn(),
    });
    const spawned: string[][] = [];
    const coordinator = loadCoordinator({ real: true, context: {
        spawn: (_command: string, args: string[]) => { spawned.push(args); return child; },
        readline: { createInterface: () => lines },
    } });
    coordinator.starts.set(5151, started);
    const worker = new coordinator.OracleWorker(40);
    const hostKills = () => coordinator.cleanup.filter((step) => step.startsWith('host:')).map((step) => Number(step.slice(5)));
    return { ...coordinator, child, lines, spawned, worker, hostKills };
}

const started = '638963100000000120';
const record = { pid: 5151, image: 'EXCEL', started };

describe('VBE oracle host process ownership', () => {
    it.each(['5151 EXCEL', '5151 EXCEL not-ticks', '0 EXCEL 100', '-1 EXCEL 100',
        '5151 EXCEL 100 extra', '5151junk EXCEL 100', 'closed'])('rejects an incomplete or invalid identity: %s', (text) => {
        expect(setup().parseRecordedProcess(text)).toBeUndefined();
    });

    it('follows the host the worker announces between cases, and forgets one it closed', () => {
        const probe = setup();
        probe.worker.start();
        probe.lines.emit('line', `XLIDE-ORACLE-HOST 5151 EXCEL ${started}`);
        expect(probe.worker.hostProcess).toEqual(record);
        probe.lines.emit('line', 'XLIDE-ORACLE-HOST closed');
        expect(probe.worker.hostProcess).toBeUndefined();
        // A worker that then crashes leaves no host of its own to kill.
        probe.child.emit('close', 1);
        expect(probe.hostKills()).toEqual([]);
    });

    it.each([0, 1])('forgets the host when the worker exits with code %s, killing it only after a crash', (code) => {
        const probe = setup();
        probe.worker.start();
        probe.lines.emit('line', `XLIDE-ORACLE-HOST 5151 EXCEL ${started}`);
        probe.child.emit('close', code);
        expect(probe.worker.hostProcess).toBeUndefined();
        expect(probe.hostKills()).toEqual(code ? [5151] : []);
    });

    it.each(['lookup', 'kill'])('keeps the host when a worker crashes and the %s that would end it is cut short', (cut) => {
        const probe = setup();
        probe.worker.start();
        probe.lines.emit('line', `XLIDE-ORACLE-HOST 5151 EXCEL ${started}`);
        if (cut === 'lookup') { probe.starts.set(5151, 'cut'); } else { probe.state.killCut = true; }
        probe.child.emit('close', 1);
        expect(probe.worker.hostProcess).toEqual(record);
    });

    it('tries again a host the worker could not end when it closed with an error', async () => {
        const probe = setup();
        probe.worker.start();
        probe.lines.emit('line', `XLIDE-ORACLE-HOST 5151 EXCEL ${started}`);
        const closing = probe.worker.close();
        // The kill after the worker's exit is cut short; the retry ends the host.
        probe.state.lookupsCut = 1;
        probe.child.emit('close', 1);
        await closing;
        expect(probe.hostKills()).toEqual([5151]);
        expect(probe.worker.hostProcess).toBeUndefined();
    });

    it('forgets the host after a clean close, so a later clean-up cannot kill a reused pid', async () => {
        const probe = setup();
        probe.worker.start();
        probe.worker.hostProcess = record;
        const closing = probe.worker.close();
        probe.child.emit('close', 0);
        await closing;
        expect(probe.worker.hostProcess).toBeUndefined();
        expect(probe.hostKills()).toEqual([]);
    });

    it.each([started, '638963100000000130', null, 'unreadable'])(
        'kills the host on a forced close only while its start time is still %s', async (current) => {
            vi.useFakeTimers();
            const probe = setup();
            probe.worker.start();
            probe.worker.hostProcess = record;
            probe.starts.set(5151, current);
            const closing = probe.worker.close();
            await vi.advanceTimersByTimeAsync(30_000);
            await closing;
            expect(probe.cleanup).toEqual(['host:4343', ...(current === started ? ['host:5151'] : [])]);
            // One still there whose start time cannot be read stays recorded.
            expect(probe.worker.hostProcess).toEqual(current === 'unreadable' ? record : undefined);
        });

    it.each(Object.entries(OFFICE_HOST_APPS).flatMap(([host, info]) =>
        [started, '638963100000000130', null].map((current) => ({ host, image: info.processName, current }))))(
        'kills the $host host a timed-out case recorded only while its start time is still the recorded one ($current)', async ({ host, image, current }) => {
            const probe = setup();
            probe.starts.set(5151, current);
            probe.worker.hostProcess = record;
            probe.worker.kill = vi.fn();
            probe.worker.request = async ({ pidPath }: { pidPath: string }) => {
                fs.writeFileSync(pidPath, `5151 ${image} ${started}`);
                return { timedOut: true };
            };
            const result = await probe.runCaseOnce(probe.worker, { id: 'probe', host }, 1, 0);
            expect(result.outcome).toBe('timeout');
            expect(probe.worker.kill).toHaveBeenCalledOnce();
            expect(probe.hostKills()).toEqual(current === started ? [5151] : []);
            expect(probe.worker.hostProcess).toBeUndefined();
            expect(probe.worker.pidPath).toBeNull();
        });

    it.each(['timedOut', 'exited'] as const)('kills the announced host when a case %s before recording a new one', async (ending) => {
        const probe = setup();
        probe.worker.kill = vi.fn();
        // The previous case's host was announced, and quitting it never returned.
        probe.worker.hostProcess = record;
        probe.worker.request = async () => (ending === 'timedOut' ? { timedOut: true } : { exited: true, code: 1 });
        const result = await probe.runCaseOnce(probe.worker, { id: 'probe' }, 1, 0);
        expect(result.outcome).toBe(ending === 'timedOut' ? 'timeout' : 'worker_error');
        expect(probe.hostKills()).toEqual([5151]);
        expect(probe.worker.hostProcess).toBeUndefined();
    });

    it.each([
        ...(['timedOut', 'exited'] as const).flatMap((ending) => [{ ending, cut: 'lookup' }, { ending, cut: 'kill' }]),
    ])('keeps the host when a case $ending and the $cut that would end it is cut short, for the clean-up to end it', async ({ ending, cut }) => {
        const probe = setup();
        probe.worker.kill = vi.fn();
        probe.worker.hostProcess = record;
        if (cut === 'lookup') { probe.starts.set(5151, 'cut'); } else { probe.state.killCut = true; }
        probe.worker.request = async () => (ending === 'timedOut' ? { timedOut: true } : { exited: true, code: 1 });
        await probe.runCaseOnce(probe.worker, { id: 'probe' }, 1, 0);
        expect(probe.worker.hostProcess).toEqual(record);
        // The Ctrl-C clean-up that follows, no longer cut short, ends it.
        probe.starts.set(5151, started);
        probe.state.killCut = false;
        probe.cleanup.length = 0;
        expect(probe.killProcess(probe.worker.hostProcess)).toBe(true);
        expect(probe.hostKills()).toEqual([5151]);
    });

    it('ends a host kept from a worker that already ended when the run closes it', async () => {
        const probe = setup();
        probe.worker.hostProcess = record;
        await probe.worker.close();
        expect(probe.hostKills()).toEqual([5151]);
        expect(probe.worker.hostProcess).toBeUndefined();
    });

    it('ends a kept host before the next case starts a worker', () => {
        const probe = setup();
        probe.worker.hostProcess = record;
        void probe.worker.request({}, 60_000);
        expect(probe.hostKills()).toEqual([5151]);
        expect(probe.spawned).toHaveLength(1);
        expect(probe.worker.hostProcess).toBeUndefined();
    });

    it('fails the next case rather than start a worker over a kept host it still cannot end', async () => {
        const probe = setup();
        probe.worker.hostProcess = record;
        probe.starts.set(5151, 'unreadable');
        const result = await probe.runCaseOnce(probe.worker, { id: 'probe' }, 1, 0);
        expect(result).toMatchObject({ outcome: 'worker_error', message: 'Office process 5151 from an earlier case could not be ended.' });
        expect(probe.spawned).toHaveLength(0);
        expect(probe.worker.hostProcess).toEqual(record);
        // Once it can be read, closing the run ends it.
        probe.starts.set(5151, started);
        await probe.worker.close();
        expect(probe.hostKills()).toEqual([5151]);
    });

    it('keeps a host whose start time cannot be read', () => {
        const probe = setup();
        probe.starts.set(5151, 'unreadable');
        expect(probe.killProcess(record)).toBe(false);
        expect(probe.hostKills()).toEqual([]);
    });

    it('forgets the previous host when a case that replied recorded none', async () => {
        const probe = setup();
        probe.worker.hostProcess = record;
        probe.worker.request = async () => ({ line: '{"outcome":"accepted"}' });
        await probe.runCaseOnce(probe.worker, { id: 'probe' }, 1, 0);
        expect(probe.worker.hostProcess).toBeUndefined();
    });

    it('knows the current case folder while the case runs, for clean-up after Ctrl-C', async () => {
        const probe = setup();
        let during: string | null = null;
        probe.worker.request = async ({ pidPath }: { pidPath: string }) => {
            during = probe.worker.pidPath;
            expect(pidPath).toBe(during);
            return { line: '{"outcome":"accepted"}' };
        };
        await probe.runCaseOnce(probe.worker, { id: 'probe' }, 1, 0);
        expect(path.basename(path.dirname(during!))).toMatch(/^xlide-vbe-oracle-/);
        expect(probe.worker.pidPath).toBeNull();
    });

    it('writes and retires the same start time in the PowerShell worker', () => {
        // A Windows checkout has CRLF lines; the snippets below are matched with LF.
        const worker = fs.readFileSync(path.join(oracleDir, 'vbe_oracle_worker.ps1'), 'utf8').replaceAll('\r\n', '\n');
        const between = (from: string, to: string) => worker.slice(worker.indexOf(from), worker.indexOf(to));
        const start = between('function Start-OracleHost', 'function Get-OwnedOracleHostProcess');
        const ticks = start.indexOf('(Get-Process -Id $processId -ErrorAction Stop).StartTime.ToUniversalTime().Ticks');
        // The pid is recorded only once its start time is read.
        expect(ticks).toBeGreaterThan(-1);
        expect(ticks).toBeLessThan(start.indexOf('$script:appPid = $processId'));
        expect(start).toContain('$script:appStarted = $ticks - ($ticks % 10)');
        expect(start.indexOf('Write-OracleHostProcess')).toBeGreaterThan(start.indexOf('$script:appStarted ='));
        const write = between('function Write-OracleHostProcess', 'function Start-OracleHost');
        expect(write).toContain('$record = "$($script:appPid) $($script:hostImageOfApp) $($script:appStarted)"');
        expect(write).toContain('Set-Content -LiteralPath $PidPath -Value $record -Encoding ascii');
        expect(write).toContain('[Console]::Out.WriteLine("XLIDE-ORACLE-HOST " + $record)');
        const owned = between('function Get-OwnedOracleHostProcess', 'function Stop-OracleHost');
        expect(owned).toContain('($ticks - ($ticks % 10)) -eq $script:appStarted');
        expect(owned).toContain('if ($null -eq $p -or $p.ProcessName -ne $script:hostImageOfApp) { return }');
        // The pid and image alone never decide.
        expect(owned.match(/return \$p\b/g)).toHaveLength(1);
        const stop = between('function Stop-OracleHost', 'function Set-OracleHostSettings');
        // Only a host whose start time matched is killed; one that cannot be read is waited for.
        // The handle is opened before the start time is read and kept until the kill.
        expect(owned).toContain("try {\n        $null = $p.Handle\n        $ticks = $p.StartTime.ToUniversalTime().Ticks\n    } catch { return 'unverified' }");
        expect(stop).toContain('Get-OwnedOracleHostProcess | Where-Object { $_ -is [Diagnostics.Process] } |\n            Stop-Process');
        expect(stop).toContain('$null -ne (Get-OwnedOracleHostProcess)');
        // A host that would not stop stays recorded: no "closed", and the pid file is kept;
        // the worker ends with an error, so its callers cannot start another over it.
        const kept = stop.indexOf('if ($null -ne (Get-OwnedOracleHostProcess)) {\n            throw ');
        expect(kept).toBeGreaterThan(stop.indexOf('Stop-Process -Force'));
        expect(kept).toBeLessThan(stop.indexOf('$script:appPid = 0'));
        expect(kept).toBeLessThan(stop.indexOf('Remove-Item -LiteralPath $PidPath'));
        expect(kept).toBeLessThan(stop.indexOf('XLIDE-ORACLE-HOST closed'));
        expect(stop).toContain('Remove-Item -LiteralPath $PidPath');
        expect(stop).toContain('[Console]::Out.WriteLine("XLIDE-ORACLE-HOST closed")');
        // A host kept from the previous case is recorded for this one too.
        const invoke = worker.slice(worker.indexOf('function Invoke-OracleCase'));
        expect(invoke).toContain('Write-OracleHostProcess\n            try { Set-OracleHostSettings $script:app } catch { Stop-OracleHost }');
    });
});
