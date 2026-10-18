import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { spawn, spawnSync } from 'node:child_process';
import { afterEach, describe, expect, it } from 'vitest';
import { fixture, loadCoordinator, removeTemporary } from './helpers/vbeOracleCoordinator';

afterEach(removeTemporary);

const otherOwner = (file: string, owner: Record<string, unknown>) => {
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, JSON.stringify(owner));
};

// A second coordinator run by the same user: another process, the same profile.
function contender(holder: ReturnType<typeof loadCoordinator>) {
    const other = loadCoordinator({ home: holder.dir });
    other.fakeProcess.pid = 4343;
    other.starts.set(4343, '200');
    return other;
}

describe('VBE oracle per-user lock', () => {
    it('refuses a second run while the first holds the lock, measuring nothing', async () => {
        const holder = loadCoordinator();
        const release = holder.acquireOracleLock();
        const other = contender(holder);
        expect((await other.run([fixture()])).exit).toBe(2);
        expect(other.errors.join('\n')).toContain('Another VBE oracle run is active (pid 4242)');
        expect(other.ran).toEqual([]);
        release();
        expect(fs.existsSync(holder.oracleLockPath())).toBe(false);
    });

    it('records this machine, pid and start time, and holds the lock while measuring', async () => {
        const run = loadCoordinator();
        run.state.measure = async () => {
            const files = fs.readdirSync(run.oracleLockPath());
            expect(files).toHaveLength(1);
            expect(JSON.parse(fs.readFileSync(path.join(run.oracleLockPath(), files[0]), 'utf8')))
                .toEqual({ host: os.hostname(), pid: 4242, started: '100' });
            const other = contender(run);
            expect((await other.run([fixture()])).exit).toBe(2);
        };
        expect((await run.run([fixture()])).exit).toBe(0);
        expect(run.ran).toHaveLength(1);
        expect(fs.existsSync(run.oracleLockPath())).toBe(false);
        expect(run.fakeProcess.eventNames()).toEqual([]);
    });

    it.each([null, '101'])('takes over the lock of a run that ended, whose pid now has start time %s', async (started) => {
        const holder = loadCoordinator();
        holder.acquireOracleLock();
        const other = contender(holder);
        other.starts.set(4242, started);
        expect((await other.run([fixture()])).exit).toBe(0);
        expect(other.ran).toHaveLength(1);
        expect(fs.existsSync(holder.oracleLockPath())).toBe(false);
    });

    it('takes over a lock whose owner runs elevated only once it has ended', () => {
        const holder = loadCoordinator();
        holder.acquireOracleLock();
        const run = contender(holder);
        // processStart reads an elevated owner through CIM in the same PowerShell.
        expect(() => run.acquireOracleLock()).toThrow('Another VBE oracle run is active (pid 4242)');
        const lookups = run.spawns.filter((text) => text.includes('Get-Process -Id 4242'));
        expect(lookups).toHaveLength(1);
        expect(lookups[0]).toContain('Get-CimInstance Win32_Process -Filter "ProcessId=4242"');
    });

    it('refuses an owner that runs but whose start time cannot be read', () => {
        const holder = loadCoordinator();
        holder.acquireOracleLock();
        const run = contender(holder);
        run.starts.set(4242, 'unreadable');
        expect(() => run.acquireOracleLock()).toThrow('Cannot tell whether process 4242, which holds the oracle lock, is an oracle run');
        expect(fs.readdirSync(run.oracleLockPath())).toHaveLength(1);
    });

    it('refuses a lock owned on another machine, whose processes it cannot see', () => {
        const run = loadCoordinator();
        const owner = path.join(run.oracleLockPath(), '00000000-0000-0000-0000-000000000000.json');
        otherOwner(owner, { host: 'elsewhere', pid: 9999, started: '1' });
        expect(() => run.acquireOracleLock()).toThrow(`An oracle run on elsewhere holds the lock; remove ${run.oracleLockPath()} only after`);
        expect(fs.existsSync(owner)).toBe(true);
    });

    it.each([
        ['no machine', { pid: 9999, started: '1' }],
        ['pid 0', { host: os.hostname(), pid: 0, started: '1' }],
        ['a pid in text', { host: os.hostname(), pid: '9999', started: '1' }],
        ['a start time not in ticks', { host: os.hostname(), pid: 9999, started: 'yesterday' }],
        ['null', null],
    ])('refuses an owner file with %s', (_name, owner) => {
        const run = loadCoordinator();
        const file = path.join(run.oracleLockPath(), '00000000-0000-0000-0000-000000000000.json');
        otherOwner(file, owner as Record<string, unknown>);
        expect(() => run.acquireOracleLock()).toThrow('Invalid lock owner');
        expect(fs.existsSync(file)).toBe(true);
    });

    it.each([['two owner files', ['a.json', 'b.json']], ['another file', ['notes.txt']]])('refuses a lock folder with %s', (_name, names) => {
        const run = loadCoordinator();
        fs.mkdirSync(run.oracleLockPath());
        for (const name of names) { fs.writeFileSync(path.join(run.oracleLockPath(), name), '{}'); }
        expect(() => run.acquireOracleLock()).toThrow('is being taken or given up by another run, or was left part-written');
    });

    it('waits for an owner file still being written, then checks its owner', () => {
        const holder = loadCoordinator();
        holder.acquireOracleLock();
        const run = contender(holder);
        const readFileSync = run.fileSystem.readFileSync;
        let reads = 0;
        run.fileSystem.readFileSync = ((...args: Parameters<typeof fs.readFileSync>) => (++reads < 3 ? '{"host":' : readFileSync(...args))) as typeof fs.readFileSync;
        expect(() => run.acquireOracleLock()).toThrow('Another VBE oracle run is active (pid 4242)');
        expect(reads).toBe(3);
    });

    it('refuses an owner file that stays partly written', () => {
        const run = loadCoordinator();
        otherOwner(path.join(run.oracleLockPath(), 'abc.json'), {});
        fs.writeFileSync(path.join(run.oracleLockPath(), 'abc.json'), '{"host":');
        expect(() => run.acquireOracleLock()).toThrow('is being taken or given up by another run, or was left part-written');
        expect(fs.existsSync(path.join(run.oracleLockPath(), 'abc.json'))).toBe(true);
    });

    it('removes an empty lock folder, which holds no run, and takes the lock', () => {
        const run = loadCoordinator();
        fs.mkdirSync(run.oracleLockPath());
        const release = run.acquireOracleLock();
        expect(fs.readdirSync(run.oracleLockPath())).toHaveLength(1);
        release();
        expect(fs.existsSync(run.oracleLockPath())).toBe(false);
    });

    it('leaves a lock folder alone when another run writes its owner file just before the empty folder is removed', () => {
        const run = loadCoordinator();
        const lock = run.oracleLockPath();
        fs.mkdirSync(lock);
        const competitor = path.join(lock, 'ffffffff-ffff-ffff-ffff-ffffffffffff.json');
        const rmdirSync = run.fileSystem.rmdirSync;
        run.fileSystem.rmdirSync = ((target: fs.PathLike) => {
            if (!fs.existsSync(competitor)) { fs.writeFileSync(competitor, JSON.stringify({ host: os.hostname(), pid: 5252, started: '400' })); }
            return rmdirSync(target);
        }) as typeof fs.rmdirSync;
        expect(() => run.acquireOracleLock()).toThrow('Another VBE oracle run is active (pid 5252)');
        expect(fs.existsSync(competitor)).toBe(true);
    });

    it('tries again when another run removes its new folder as empty before its owner file is in it', () => {
        const run = loadCoordinator();
        const lock = run.oracleLockPath();
        const writeFileSync = run.fileSystem.writeFileSync;
        let removed = false;
        run.fileSystem.writeFileSync = ((target: fs.PathOrFileDescriptor, ...rest: unknown[]) => {
            if (!removed && path.dirname(String(target)) === lock) { removed = true; fs.rmdirSync(lock); }
            return (writeFileSync as (...args: unknown[]) => unknown)(target, ...rest);
        }) as typeof fs.writeFileSync;
        const release = run.acquireOracleLock();
        expect(removed).toBe(true);
        expect(fs.readdirSync(lock)).toHaveLength(1);
        release();
    });

    it('backs off when another run took the folder over while it wrote its owner file', () => {
        const run = loadCoordinator();
        const lock = run.oracleLockPath();
        const competitor = path.join(lock, 'ffffffff-ffff-ffff-ffff-ffffffffffff.json');
        const writeFileSync = run.fileSystem.writeFileSync;
        run.fileSystem.writeFileSync = ((target: fs.PathOrFileDescriptor, ...rest: unknown[]) => {
            const written = (writeFileSync as (...args: unknown[]) => unknown)(target, ...rest);
            if (String(target) !== competitor && path.dirname(String(target)) === lock) {
                fs.writeFileSync(competitor, JSON.stringify({ host: os.hostname(), pid: 5252, started: '400' }));
            }
            return written;
        }) as typeof fs.writeFileSync;
        expect(() => run.acquireOracleLock()).toThrow('The oracle lock changed while acquiring it');
        expect(fs.readdirSync(lock)).toEqual([path.basename(competitor)]);
    });

    it('tries again when the lock is released between seeing it and reading it', () => {
        const run = loadCoordinator();
        const lock = run.oracleLockPath();
        fs.mkdirSync(lock);
        const readdirSync = run.fileSystem.readdirSync;
        let released = false;
        run.fileSystem.readdirSync = ((target: fs.PathLike, ...rest: unknown[]) => {
            if (!released) { released = true; fs.rmdirSync(lock); }
            return (readdirSync as (...args: unknown[]) => unknown)(target, ...rest);
        }) as typeof fs.readdirSync;
        const release = run.acquireOracleLock();
        expect(released).toBe(true);
        release();
    });

    it('never removes a lock another run took over first from the same ended owner', () => {
        const holder = loadCoordinator();
        holder.acquireOracleLock();
        const run = contender(holder);
        run.starts.set(4242, null);
        const lock = run.oracleLockPath();
        const competitor = path.join(lock, 'ffffffff-ffff-ffff-ffff-ffffffffffff.json');
        const unlinkSync = run.fileSystem.unlinkSync;
        run.fileSystem.unlinkSync = ((target: fs.PathLike) => {
            if (!fs.existsSync(competitor)) {
                // The other taker removed the old owner and wrote its own first.
                unlinkSync(target);
                fs.writeFileSync(competitor, JSON.stringify({ host: os.hostname(), pid: 5252, started: '400' }));
            }
            return unlinkSync(target);
        }) as typeof fs.unlinkSync;
        expect(() => run.acquireOracleLock()).toThrow('Another VBE oracle run is active (pid 5252)');
        expect(fs.existsSync(competitor)).toBe(true);
    });

    it('does not let an ended owner release the run that took its lock over', () => {
        const holder = loadCoordinator();
        const oldRelease = holder.acquireOracleLock();
        const run = contender(holder);
        run.starts.set(4242, null);
        const release = run.acquireOracleLock();
        oldRelease();
        expect(fs.readdirSync(run.oracleLockPath())).toHaveLength(1);
        release();
        expect(fs.existsSync(run.oracleLockPath())).toBe(false);
    });

    it('refuses when its own start time cannot be read', () => {
        const run = loadCoordinator();
        run.starts.set(4242, 'unreadable');
        expect(() => run.acquireOracleLock()).toThrow('Cannot read this process start time.');
        expect(fs.existsSync(run.oracleLockPath())).toBe(false);
    });

    it('keeps the lock in the account profile, whatever HOME or USERPROFILE says', async () => {
        const holder = loadCoordinator();
        const release = holder.acquireOracleLock();
        const other = loadCoordinator({ home: holder.dir, envHome: path.join(holder.dir, 'elsewhere') });
        other.fakeProcess.pid = 4343;
        other.starts.set(4343, '200');
        expect(other.oracleLockPath()).toBe(path.join(holder.dir, '.xlide-vbe-oracle.lock'));
        expect((await other.run([fixture()])).exit).toBe(2);
        release();
        const noAccount = loadCoordinator({ envHome: path.join(holder.dir, 'home') });
        noAccount.state.account = false;
        expect(noAccount.oracleLockPath()).toBe(path.join(holder.dir, 'home', '.xlide-vbe-oracle.lock'));
    });
});

describe('VBE oracle process start times', () => {
    it.each([
        ['a running process', '638963100000000120\r\n', '638963100000000120'],
        ['an ended one', 'gone\r\n', null],
        ['one neither lookup can read', 'unreadable\r\n', 'unreadable'],
    ])('reads %s in one PowerShell', (_name, stdout, expected) => {
        const calls: string[][] = [];
        const run = loadCoordinator({ context: { spawnSync: (_command: string, args: string[]) => { calls.push(args); return { status: 0, stdout }; } } });
        expect(run.processStart(4242)).toBe(expected);
        expect(calls).toHaveLength(1);
        const script = calls[0][4];
        expect(script).toContain('$time = $null');
        expect(script).toContain('try { $time = $p.StartTime } catch { }');
        // An elevated process's StartTime reads as nothing, without an error.
        expect(script).toContain('if ($null -eq $time) {');
        expect(script).toContain('Get-CimInstance Win32_Process -Filter "ProcessId=4242"');
        expect(script).toContain('$ticks - ($ticks % 10)');
    });

    it.each([
        { status: 2, stderr: 'access denied' },
        { status: null, error: new Error('spawnSync powershell.exe ETIMEDOUT') },
        { status: 0, stdout: 'unexpected output' },
        { status: 1, stdout: 'gone' },
        { status: null, stdout: '638963100000000120', error: new Error('spawnSync powershell.exe ETIMEDOUT') },
    ])('never takes a failed lookup for an ended process: %j', (completed) => {
        const run = loadCoordinator({ context: { spawnSync: () => completed } });
        expect(() => run.processStart(4242)).toThrow('Cannot check process 4242:');
    });
});

describe('VBE oracle host kills', () => {
    it('checks and kills the host under one handle, so its pid cannot change hands in between', () => {
        const calls: string[][] = [];
        const run = loadCoordinator({ context: { spawnSync: (command: string, args: string[]) => {
            calls.push([command, ...args]);
            return { status: 0, stdout: 'killed\r\n' };
        } } });
        expect(run.killProcess({ pid: 5151, image: 'EXCEL', started: '300' })).toBe(true);
        expect(calls).toHaveLength(1);
        const script = calls[0][5];
        // Another program on the pid is not the host.
        expect(script).toContain("if ($null -eq $p -or $p.ProcessName -ne 'EXCEL') { 'gone'; exit }");
        const handle = script.indexOf('$null = $p.Handle');
        expect(handle).toBeGreaterThan(script.indexOf("$p.ProcessName -ne 'EXCEL'"));
        expect(script.indexOf('$p.StartTime')).toBeGreaterThan(handle);
        expect(script.indexOf('-ne 300) {')).toBeGreaterThan(script.indexOf('$p.StartTime'));
        expect(script.indexOf('taskkill.exe /PID 5151 /T /F')).toBeGreaterThan(script.indexOf('-ne 300) {'));
        // taskkill's own message must not stop the script before it answers.
        expect(script.indexOf("$ErrorActionPreference = 'Continue'")).toBeGreaterThan(script.indexOf('-ne 300) {'));
        expect(script.indexOf("$ErrorActionPreference = 'Continue'")).toBeLessThan(script.indexOf('taskkill.exe'));
    });

    it.each([['killed', true], ['gone', true], ['unreadable', false], ['failed', false], ['', false]])(
        'reads the kill script answering %j as ended: %s', (stdout, ended) => {
            const run = loadCoordinator({ context: { spawnSync: () => ({ status: 0, stdout: `${stdout}\r\n` }) } });
            expect(run.killProcess({ pid: 5151, image: 'EXCEL', started: '300' })).toBe(ended);
        });

    it.each([{ status: null, signal: 'SIGINT', stdout: 'killed' }, { status: 1, stdout: 'killed' },
        { status: null, error: new Error('spawnSync powershell.exe ETIMEDOUT'), stdout: '' },
        { status: 0, error: new Error('spawnSync powershell.exe ETIMEDOUT'), stdout: 'killed' }])(
        'never takes a failed kill script for an ended host: %j', (completed) => {
            const run = loadCoordinator({ context: { spawnSync: () => completed } });
            expect(run.killProcess({ pid: 5151, image: 'EXCEL', started: '300' })).toBe(false);
        });
});

describe('VBE oracle process start times on Windows', () => {
    // A normal shell cannot read StartTime of an elevated process; this
    // PowerShell's Get-Process is replaced by one whose StartTime is denied,
    // by an error or, as PowerShell 5.1 does, by reading as nothing.
    it.skipIf(process.platform !== 'win32').each([['throws', "throw 'Access denied'"], ['reads as nothing', '$null']])(
        'reads the same start time through CIM when StartTime %s', (_name, denial) => {
        const real = loadCoordinator({ context: { spawnSync } });
        const normal = real.processStart(process.pid);
        expect(normal).toMatch(/^\d+0$/);
        const denied = loadCoordinator({ context: { spawnSync: (command: string, args: string[], options: object) => spawnSync(command, [...args.slice(0, 4),
            `function Get-Process { [CmdletBinding()] param([int]$Id) $p = New-Object PSObject; $p | Add-Member ScriptProperty StartTime { ${denial} }; return $p }\n${args[4]}`], options) } });
        expect(denied.processStart(process.pid)).toBe(normal);
        expect(real.processStart(2 ** 30)).toBeNull();
    }, 45_000);
});

describe('VBE oracle host kills on Windows', () => {
    // A real process: a kill for another start time leaves it running, the
    // one for its own ends it.
    it.skipIf(process.platform !== 'win32')('kills a process only while its start time is the recorded one', async () => {
        const child = spawn(process.execPath, ['-e', 'setTimeout(() => {}, 60000)'], { stdio: 'ignore' });
        try {
            const run = loadCoordinator({ context: { spawnSync } });
            const started = run.processStart(child.pid!);
            expect(started).toMatch(/^\d+0$/);
            const other = String(BigInt(started) + 10n);
            expect(run.killProcess({ pid: child.pid!, image: 'NODE', started: other })).toBe(true);
            expect(run.processStart(child.pid!)).toBe(started);
            const exited = new Promise((resolve) => child.once('exit', resolve));
            expect(run.killProcess({ pid: child.pid!, image: 'NODE', started })).toBe(true);
            await exited;
            expect(run.processStart(child.pid!)).toBeNull();
        } finally { child.kill(); }
    }, 45_000);
});

describe('VBE oracle clean-up on exit and signals', () => {
    it.each([['exit', undefined], ['SIGINT', 130], ['SIGHUP', 129], ['SIGBREAK', 149]] as const)(
        'ends the worker and both host records, removes the case folder, then releases the lock on %s', async (event, code) => {
            const run = loadCoordinator();
            const rmSync = run.fileSystem.rmSync;
            run.fileSystem.rmSync = ((file: fs.PathLike, options?: fs.RmOptions) => {
                if (path.dirname(String(file)) === run.oracleLockPath() && fs.existsSync(file)) { run.cleanup.push('release'); }
                return rmSync(file, options);
            }) as typeof fs.rmSync;
            let caseFolder = '';
            run.state.measure = async (worker) => {
                worker.pidPath = run.caseFolder('5252 WINWORD 400');
                caseFolder = path.dirname(worker.pidPath);
                if (code === undefined) { run.fakeProcess.emit(event, 1); }
                else {
                    expect(() => run.fakeProcess.emit(event)).toThrow(`exit ${code}`);
                    // Still listening: a second Ctrl-C waits for the clean-up rather than ending it.
                    expect(run.fakeProcess.listenerCount(event)).toBe(1);
                    expect(() => run.fakeProcess.emit(event)).toThrow(`exit ${code}`);
                }
                expect(fs.existsSync(run.oracleLockPath())).toBe(false);
                // Exit follows a signal; the clean-up runs once.
                run.fakeProcess.emit('exit', code ?? 1);
            };
            await run.run([fixture()]);
            expect(run.cleanup).toEqual(['worker', 'host:5151', 'host:5252', 'release']);
            expect(fs.existsSync(caseFolder)).toBe(false);
            expect(fs.existsSync(run.dir)).toBe(true);
        });

    it.each([false, true])('names a host it could not end when the run ends: %s', async (keepHost) => {
        const run = loadCoordinator();
        run.state.keepHost = keepHost;
        await run.run([fixture()]);
        expect(run.errors.filter((line) => line.startsWith('Could not end Office process')))
            .toEqual(keepHost ? ['Could not end Office process 5151 (EXCEL); close it by hand.'] : []);
    });

    it('checks a host once when the case record names the announced one', async () => {
        const run = loadCoordinator();
        run.state.measure = async (worker) => {
            worker.pidPath = run.caseFolder('5151 EXCEL 300');
            run.spawns.length = 0;
            run.fakeProcess.emit('exit', 1);
        };
        await run.run([fixture()]);
        expect(run.cleanup).toEqual(['worker', 'host:5151']);
        expect(run.spawns.filter((script) => script.includes('Get-Process -Id 5151'))).toHaveLength(1);
    });

    it('still ends the recorded host when it reuses the announced pid with another start time', async () => {
        const run = loadCoordinator();
        run.state.measure = async (worker) => {
            // The announced host has ended and its pid now belongs to the case's own host.
            worker.pidPath = run.caseFolder('5151 EXCEL 301');
            run.starts.set(5151, '301');
            run.fakeProcess.emit('exit', 1);
        };
        await run.run([fixture()]);
        expect(run.cleanup).toEqual(['worker', 'host:5151']);
    });

    it.each(['announced', 'recorded', 'both'] as const)('leaves alone a %s host whose pid was reused', async (reused) => {
        const run = loadCoordinator();
        run.state.measure = async (worker) => {
            worker.pidPath = run.caseFolder('5252 WINWORD 400');
            if (reused !== 'recorded') { run.starts.set(5151, '301'); }
            if (reused !== 'announced') { run.starts.set(5252, '401'); }
            run.fakeProcess.emit('exit', 1);
        };
        await run.run([fixture()]);
        expect(run.cleanup).toEqual(['worker', ...(reused === 'recorded' ? ['host:5151'] : []),
            ...(reused === 'announced' ? ['host:5252'] : [])]);
    });

    it('goes on with the rest of the clean-up when one step fails', async () => {
        const run = loadCoordinator();
        run.state.measure = async (worker) => {
            (worker as unknown as { kill: () => void }).kill = () => { throw new Error('taskkill failed'); };
            run.fakeProcess.emit('exit', 1);
        };
        await run.run([fixture()]);
        expect(run.cleanup).toEqual(['host:5151']);
        expect(fs.existsSync(run.oracleLockPath())).toBe(false);
    });

    it('installs the handlers before it takes the lock', async () => {
        const run = loadCoordinator();
        let listening: number | undefined;
        const mkdirSync = run.fileSystem.mkdirSync;
        run.fileSystem.mkdirSync = ((target: fs.PathLike, ...rest: unknown[]) => {
            if (String(target) === run.oracleLockPath()) { listening ??= run.fakeProcess.listenerCount('SIGINT'); }
            return (mkdirSync as (...args: unknown[]) => unknown)(target, ...rest);
        }) as typeof fs.mkdirSync;
        expect((await run.run([fixture()])).exit).toBe(0);
        expect(listening).toBe(1);
    });

    it('releases a lock taken while a signal waited, once the signal is handled', async () => {
        const run = loadCoordinator();
        const mkdirSync = run.fileSystem.mkdirSync;
        let handled: unknown;
        run.fileSystem.mkdirSync = ((target: fs.PathLike, ...rest: unknown[]) => {
            const made = (mkdirSync as (...args: unknown[]) => unknown)(target, ...rest);
            // Node delivers a signal once the code taking the lock has finished.
            if (String(target) === run.oracleLockPath()) {
                setImmediate(() => { try { run.fakeProcess.emit('SIGINT'); } catch (err) { handled = err; } });
            }
            return made;
        }) as typeof fs.mkdirSync;
        run.state.measure = () => new Promise((resolve) => { setImmediate(resolve); });
        await run.run([fixture()]);
        expect(String(handled)).toContain('exit 130');
        expect(fs.existsSync(run.oracleLockPath())).toBe(false);
    });

    it.each([
        ['an early error', (run: ReturnType<typeof loadCoordinator>) => { run.state.build = null; }, ['--checkpoint', 'x.jsonl'], 2],
        ['an oracle failure', (run: ReturnType<typeof loadCoordinator>) => { run.state.outcome = 'worker_error'; }, [], 1],
    ] as const)('releases the lock after %s', async (_name, arrange, flags, exit) => {
        const run = loadCoordinator();
        arrange(run);
        expect((await run.run([fixture()], flags.map((flag) => (flag.endsWith('.jsonl') ? path.join(run.dir, flag) : flag)))).exit).toBe(exit);
        expect(fs.existsSync(run.oracleLockPath())).toBe(false);
        expect(run.fakeProcess.eventNames()).toEqual([]);
    });

    it.each(['measure', 'close'] as const)('releases the lock after a thrown %s error', async (stage) => {
        const run = loadCoordinator();
        if (stage === 'measure') { run.state.measure = async () => { throw new Error('measure failed'); }; }
        else { run.state.closeError = true; }
        await expect(run.run([fixture()])).rejects.toThrow(`${stage} failed`);
        expect(fs.existsSync(run.oracleLockPath())).toBe(false);
        expect(run.fakeProcess.eventNames()).toEqual([]);
    });

    it.each([
        ['an unknown argument', ['--typo'], [fixture()]],
        ['an unknown case', ['--case', 'missing'], [fixture()]],
        ['an unknown host', [], [{ ...fixture(), host: 'other' }]],
        // Names every object has are not hosts either.
        ['host constructor', [], [{ ...fixture(), host: 'constructor' }]],
        ['host __proto__', [], [{ ...fixture(), host: '__proto__' }]],
    ] as const)('refuses %s before taking the lock', async (_name, flags, cases) => {
        const run = loadCoordinator();
        expect((await run.run([...cases], [...flags])).exit).toBe(2);
        expect(run.spawns).toEqual([]);
        expect(fs.existsSync(run.oracleLockPath())).toBe(false);
    });
});

describe('VBE oracle option values', () => {
    it.each(['--case', '--cases', '--checkpoint'].flatMap((flag) => [[flag, []], [flag, ['']], [flag, ['--strict']]] as const))(
        'refuses %s with the value %j, rather than take the next option for it', (flag, rest) => {
            const run = loadCoordinator();
            expect(() => run.parseArgs([flag, ...rest])).toThrow(`${flag} expects a value`);
        });

    it('takes a value that does not start with --', () => {
        const run = loadCoordinator();
        expect(run.parseArgs(['--case', '-x', '--checkpoint', 'v.jsonl'])).toMatchObject({ caseIds: ['-x'], checkpoint: 'v.jsonl' });
    });
});
