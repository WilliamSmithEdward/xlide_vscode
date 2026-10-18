import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import crypto from 'node:crypto';
import { EventEmitter } from 'node:events';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { runInNewContext } from 'node:vm';

// Runs the real coordinator in a VM context with Windows, PowerShell and
// Office replaced, so its lock, clean-up and reporting paths run as written.
export const oracleDir = path.join(process.cwd(), 'syntax_corpus/oracle');
export const coordinatorSource = fs.readFileSync(path.join(oracleDir, 'run_vbe_oracle.mjs'), 'utf8');
export const build = '16.0.20430.20182';
export type Fixture = Record<string, unknown> & { id: string };
export type FakeWorker = { pidPath: string | null; hostProcess?: unknown };

const temporary: string[] = [];

export function removeTemporary(): void {
    for (const dir of temporary.splice(0)) { fs.rmSync(dir, { recursive: true, force: true }); }
}

export function fixture(id = 'probe'): Fixture {
    return { id, host: 'excel', mode: 'compile', source: 'Sub Probe()\nEnd Sub', expected: 'rejected' };
}

export type CoordinatorOptions = {
    // The account's profile folder, where the per-user lock lives.
    home?: string;
    // What HOME or USERPROFILE says, which the lock must not follow.
    envHome?: string;
    // Keep the real worker, registration and measurement; tests then pass their own spawn.
    real?: boolean;
    context?: Record<string, unknown>;
};

export function loadCoordinator(options: CoordinatorOptions = {}) {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-oracle-test-'));
    temporary.push(dir);
    const home = options.home ?? dir;
    const state = {
        build: build as string | null, outcome: 'rejected', stage: 'compile_dialog', hostname: os.hostname(), account: true,
        measure: undefined as ((worker: FakeWorker) => Promise<void>) | undefined, closed: 0, closeError: false, killCut: false, lookupsCut: 0, keepHost: false,
    };
    // Each pid's start time as processStart reports it: null once it has ended,
    // and 'cut' for a lookup cut short, as by a Ctrl-C.
    const starts = new Map<number, string | null>([[4242, '100'], [5151, '300'], [5252, '400']]);
    const spawns: string[] = [];
    // The clean-up steps in order: the worker, then each host killed.
    const cleanup: string[] = [];
    const output: string[] = [];
    const errors: string[] = [];
    const ran: Fixture[] = [];
    const fakeProcess = Object.assign(new EventEmitter(), {
        platform: 'win32', pid: 4242, env: {} as Record<string, string>,
        exit: (code: number) => {
            fakeProcess.emit('exit', code);
            throw new Error(`exit ${code}`);
        },
    });
    const fileSystem = { ...fs };
    const context: Record<string, unknown> = {
        fs: fileSystem, path, crypto, fileURLToPath, pathToFileURL, setTimeout, clearTimeout, Buffer,
        os: {
            ...os, tmpdir: () => dir, hostname: () => state.hostname, homedir: () => options.envHome ?? home,
            userInfo: () => {
                if (!state.account) { throw Object.assign(new Error('no account record'), { code: 'ENOENT' }); }
                return { ...os.userInfo(), homedir: home };
            },
        },
        coordinatorUrl: pathToFileURL(path.join(oracleDir, 'run_vbe_oracle.mjs')).href,
        spawnSync: (command: string, args: string[]) => {
            if (command === 'taskkill') {
                cleanup.push(`host:${args[1]}`);
                return { status: 0 };
            }
            spawns.push(args[4]);
            const pid = /Get-Process -Id (\d+)/.exec(args[4])?.[1];
            // killProcess: checks the start time and kills in one PowerShell.
            const expected = /-ne (\d+)\) \{ 'gone'; exit \}/.exec(args[4])?.[1];
            if (pid && expected) {
                const current = starts.get(Number(pid));
                if (current === 'cut' || (state.lookupsCut > 0 && state.lookupsCut--)) { return { status: null, signal: 'SIGINT', stdout: '' }; }
                if (current === 'unreadable') { return { status: 0, stdout: 'unreadable\r\n' }; }
                if (current !== expected) { return { status: 0, stdout: 'gone\r\n' }; }
                cleanup.push(`host:${pid}`);
                return state.killCut ? { status: null, signal: 'SIGINT', stdout: '' } : { status: 0, stdout: 'killed\r\n' };
            }
            if (pid && (starts.get(Number(pid)) === 'cut' || (state.lookupsCut > 0 && state.lookupsCut--))) {
                return { status: null, signal: 'SIGINT', stdout: '' };
            }
            return { status: 0, stdout: pid ? `${starts.get(Number(pid)) ?? 'gone'}\r\n` : '' };
        },
        process: fakeProcess,
        console: { log: (line: string) => output.push(line), error: (line: string) => errors.push(line) },
        probeBuild: () => ({ registered: true, build: state.build, why: 'unreadable' }),
        FakeWorker: class {
            hostProcess: unknown = { pid: 5151, image: 'EXCEL', started: '300' };
            pidPath: string | null = null;
            async close() {
                state.closed++;
                if (state.closeError) { throw new Error('close failed'); }
                // Like the real one, it forgets a host it ended.
                if (!state.keepHost) { this.hostProcess = undefined; }
            }
            kill() { cleanup.push('worker'); }
        },
        measureCase: async (worker: FakeWorker, kase: Fixture) => {
            ran.push(kase);
            await state.measure?.(worker);
            return { caseId: kase.id, outcome: state.outcome, stage: state.stage, hostBuild: state.build };
        },
        ...options.context,
    };
    const api = runInNewContext(`${coordinatorSource.replace(/^import .*;\r?\n/gm, '')
        .replaceAll('import.meta.url', 'coordinatorUrl')
        .replace(/^process\.exit\(await .*\);$/m, '')}
        ${options.real ? '' : `registeredHostBuild = probeBuild;
        comRegistered = () => true;
        OracleWorker = FakeWorker;
        runCaseWithRetries = measureCase;`}
        ({ main, parseArgs, acquireOracleLock, oracleLockPath, processStart, killProcess, parseRecordedProcess,
            OracleWorker, runCaseOnce })`, context);
    const casesFile = path.join(dir, 'cases.json');
    async function run(cases: Fixture[], flags: string[] = []) {
        fs.writeFileSync(casesFile, JSON.stringify({ cases }));
        output.length = 0;
        errors.length = 0;
        const exit = await api.main(['--cases', casesFile, '--json', ...flags]);
        return { exit, report: output.length ? JSON.parse(output[0]) : undefined };
    }
    // A case folder as the coordinator makes one, holding the host record a worker wrote.
    function caseFolder(record?: string) {
        const folder = fs.mkdtempSync(path.join(dir, 'xlide-vbe-oracle-'));
        if (record) { fs.writeFileSync(path.join(folder, 'host.pid'), record); }
        return path.join(folder, 'host.pid');
    }
    return { ...api, dir, casesFile, state, starts, spawns, cleanup, output, errors, ran, fakeProcess, fileSystem, run, caseFolder };
}
