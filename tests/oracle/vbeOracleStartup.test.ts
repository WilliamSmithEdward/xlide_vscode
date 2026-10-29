import { EventEmitter } from 'node:events';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { runInNewContext } from 'node:vm';
import { describe, expect, it, vi } from 'vitest';
import { OFFICE_HOST_APPS } from '../../src/officeHostApps';

const source = readFileSync(join(process.cwd(), 'syntax_corpus/oracle/run_vbe_oracle.mjs'), 'utf8');
const names = ['comRegistered', 'registeredHostBuild', 'caseHost', 'unsupportedHost', 'malformedReferences',
    'caseBitness', 'unsupportedBitness', 'otherBitnessResult', 'hostsNeedingBuilds', 'knownHostCrashResult', 'expectedMatches', 'isOracleInfrastructureFailure'];
const functions = names.map((name) => {
    const start = source.indexOf(`function ${name}(`);
    return source.slice(start, source.indexOf('\n}', start) + 2);
}).join('\n');
const mainStart = source.indexOf('async function main(');
const mainSource = source.slice(mainStart, source.indexOf('\nprocess.exit(', mainStart));
const hosts = Object.entries(OFFICE_HOST_APPS);
const hostModes = hosts.flatMap(([host, info]) => [false, true].map((checkpoint) => ({ host, info, checkpoint })));
const build = '16.0.20430.20182';
const marker = 'XLIDE-ORACLE-REGISTERED';
const timeoutError = new Error('spawnSync powershell.exe ETIMEDOUT');
type Case = { id: string; host?: string; crashesHostBuilds?: string[] };
type Completed = { status: number | null; stdout?: string | null; stderr?: string | null; error?: Error };

function setup(cases: Case[] = [], checkpoint = false, exeBitness: (file: unknown) => number | undefined = () => 64) {
    const spawnSync = vi.fn((_command: string, _args: string[], _options: object): Completed => ({ status: 0 }));
    const errors: string[] = [];
    const output: string[] = [];
    const ran: string[] = [];
    const api = runInNewContext(`
        const DEFAULT_HOST = 'excel';
        const EVIDENCE_OUTCOMES = new Set(['accepted', 'rejected']);
        ${functions}
        ${mainSource}
        ({ main, comRegistered, registeredHostBuild })`, {
        spawnSync,
        OFFICE_HOSTS: Object.fromEntries(hosts.map(([host, info]) => [host,
            { progId: info.progId, image: info.processName, noun: info.noun }])),
        parseArgs: () => ({ timeout: 30, timeoutRetries: 0, dialogHoldSeconds: 0,
            casesPerInstance: 40, caseIds: [], json: true, checkpoint: checkpoint ? 'checkpoint' : undefined }),
        process: Object.assign(new EventEmitter(), { platform: 'win32' }),
        setDialogSounds: () => {}, // Sound lifecycle is covered in vbeOracleDialogSounds.test.ts.
        console: { log: (line: string) => output.push(line), error: (line: string) => errors.push(line) },
        loadCasesDocument: () => ({ cases }),
        loadCheckpoint: () => new Map(),
        // The run lock is covered in vbeOracleRunLock.test.ts; here it is always free.
        acquireOracleRunLock: async () => ({ held: true, release: async () => {} }),
        oracleRunLockAddress: () => 'test-lock', oracleRunOwner: () => ({ pid: 1 }), describeRefusal: () => 'refused',
        // Bitness is covered in vbeOracleBitness.test.ts; here every host is 64-bit.
        exeBitness,
        // Owned hosts are covered in vbeOracleOwnedHosts.test.ts; here the run owns none.
        OwnedHosts: class { load() { return []; } add() {} sweep() { return []; } },
        ownedHostsPath: () => 'hosts.json', describeLeftHosts: () => 'left', abandonActiveCases: () => {},
        unwritableCheckpoint: () => undefined, // A real path is checked in vbeOracleOptions.test.ts.
        OracleWorker: class { async close() {} },
        caseHash: () => 'hash', checkpointKey: () => 'key', phaseMatches: () => true,
        fs: { appendFileSync: vi.fn() },
        runCaseWithRetries: async (_worker: unknown, kase: Case) => {
            ran.push(kase.id);
            return { caseId: kase.id, outcome: 'accepted' };
        },
    });
    return { ...api, spawnSync, errors, output, ran };
}

const combinedClassifications: { name: string; completed: Completed; registered: boolean; build: string | null; why?: string }[] = [
    { name: 'unregistered', completed: { status: 2 }, registered: false, build: null, why: 'PowerShell exited 2' },
    { name: 'ENOENT without stdout', completed: { status: null, error: new Error('spawnSync powershell.exe ENOENT') },
        registered: false, build: null, why: 'spawnSync powershell.exe ENOENT' },
    { name: 'EMFILE without stdout', completed: { status: null, error: new Error('spawnSync powershell.exe EMFILE') },
        registered: false, build: null, why: 'spawnSync powershell.exe EMFILE' },
    { name: 'stdout null', completed: { status: null, stdout: null, error: new Error('spawnSync powershell.exe EMFILE') },
        registered: false, build: null, why: 'spawnSync powershell.exe EMFILE' },
    { name: 'ENOBUFS before marker', completed: { status: null, stdout: 'noise\n', error: new Error('spawnSync powershell.exe ENOBUFS') },
        registered: false, build: null, why: 'spawnSync powershell.exe ENOBUFS' },
    { name: 'timeout before marker', completed: { status: null, error: timeoutError },
        registered: false, build: null, why: timeoutError.message },
    { name: 'timeout with incomplete marker', completed: { status: null, stdout: `${marker}\r`, error: timeoutError },
        registered: false, build: null, why: timeoutError.message },
    { name: 'exit 2 before marker', completed: { status: 2, stdout: 'other output\n' },
        registered: false, build: null, why: 'PowerShell exited 2' },
    { name: 'build without marker', completed: { status: 0, stdout: `${build}\n` },
        registered: false, build: null, why: 'PowerShell exited 0' },
    { name: 'marker inside a line', completed: { status: 0, stdout: `prefix ${marker}\n${build}\n` },
        registered: false, build: null, why: 'PowerShell exited 0' },
    { name: 'unterminated marker', completed: { status: 0, stdout: marker },
        registered: false, build: null, why: 'PowerShell exited 0' },
    { name: 'marker after bare CR', completed: { status: 0, stdout: `noise\r${marker}\n${build}\n` },
        registered: false, build: null, why: 'PowerShell exited 0' },
    { name: 'marker after Unicode line separator', completed: { status: 0, stdout: `noise\u2028${marker}\n${build}\n` },
        registered: false, build: null, why: 'PowerShell exited 0' },
    { name: 'repeated markers', completed: { status: 0, stdout: `${marker}\n${build}\n${marker}\n` }, registered: true, build },
    { name: 'unterminated repeated marker', completed: { status: 0, stdout: `${marker}\n${build}\n${marker}` },
        registered: true, build: null, why: 'PowerShell exited 0' },
    { name: 'readable build LF', completed: { status: 0, stdout: `${marker}\n${build}\n` }, registered: true, build },
    { name: 'readable build CRLF', completed: { status: 0, stdout: `${marker}\r\n${build}\r\n` }, registered: true, build },
    { name: 'marker not on first line', completed: { status: 0, stdout: `${build}\n${marker}\n` }, registered: true, build },
    { name: 'noise before marker', completed: { status: 0, stdout: `noise\n${marker}\n${build}\n` },
        registered: true, build: null, why: 'PowerShell exited 0' },
    { name: 'empty build', completed: { status: 2, stdout: `${marker}\n` },
        registered: true, build: null, why: 'PowerShell exited 2' },
    { name: 'invalid build', completed: { status: 0, stdout: `${marker}\n16.0.123\n` },
        registered: true, build: null, why: 'PowerShell exited 0' },
    { name: 'five-part build', completed: { status: 0, stdout: `${marker}\n16.0.1.2.3\n` },
        registered: true, build: null, why: 'PowerShell exited 0' },
    { name: 'build without final newline', completed: { status: 0, stdout: `${marker}\n${build}` }, registered: true, build },
    { name: 'stderr build error', completed: { status: 2, stdout: `${marker}\n`, stderr: ' File version unavailable\r\n' },
        registered: true, build: null, why: 'File version unavailable' },
    { name: 'nonzero exit with build', completed: { status: 2, stdout: `${marker}\n${build}\n` },
        registered: true, build: null, why: 'PowerShell exited 2' },
    { name: 'timeout after marker', completed: { status: null, stdout: `${marker}\n${build}\n`,
        stderr: 'ignored stderr', error: timeoutError }, registered: true, build: null, why: timeoutError.message },
    { name: 'ENOBUFS after marker', completed: { status: null, stdout: `${marker}\n${build}\n`,
        error: new Error('spawnSync powershell.exe ENOBUFS') }, registered: true, build: null, why: 'spawnSync powershell.exe ENOBUFS' },
];

const classifications = combinedClassifications.flatMap((row) => row.registered
    ? [{ ...row, fallback: undefined }]
    : [false, true].map((fallback) => ({ ...row,
        name: `${row.name}; fallback ${fallback ? 'registered' : 'not registered'}`, registered: fallback, fallback })));

describe('VBE oracle synchronous startup probes', () => {
    it.each(classifications.map((row) => [row.name, row] as const))('classifies %s', (_name, { completed, registered, build: expectedBuild, why, fallback }) => {
        const probe = setup();
        probe.spawnSync.mockReturnValueOnce(completed).mockReturnValue({ status: fallback ? 0 : 2 });
        expect(probe.registeredHostBuild('Excel.Application')).toEqual({ registered, build: expectedBuild,
            ...(why ? { why } : {}) });
        expect(probe.spawnSync).toHaveBeenCalledTimes(fallback === undefined ? 1 : 2);
        if (fallback !== undefined) {
            expect(probe.spawnSync.mock.calls[1][2]).toEqual({ encoding: 'utf8', timeout: 5000 });
        }
    });

    it.each(['ENOENT', 'EMFILE', 'ETIMEDOUT', 'ENOBUFS'])('treats a fallback %s as a registration failure and preserves the build diagnostic', (code) => {
        const probe = setup();
        probe.spawnSync.mockReturnValueOnce({ status: 2, stdout: null, stderr: 'build unavailable' })
            .mockReturnValueOnce({ status: null, error: new Error(`spawnSync powershell.exe ${code}`) });
        expect(probe.registeredHostBuild('Excel.Application')).toEqual({
            registered: false, build: null, why: 'build unavailable',
        });
        expect(probe.spawnSync).toHaveBeenCalledTimes(2);
        expect(probe.spawnSync.mock.calls[1][2]).toEqual({ encoding: 'utf8', timeout: 5000 });
    });

    it.each(hosts)('uses one probe process for %s when no build is needed, for its exe', async (host, info) => {
        // Every run needs each host's exe, whose bitness decides what runs.
        const probe = setup([{ id: 'safe', host }, { id: 'also-safe', host }]);
        probe.spawnSync.mockReturnValue({ status: 0, stdout: `${marker}\r\nXLIDE-ORACLE-EXE C:\\Office\\${info.processName}.EXE\r\n${build}\r\n` });
        expect(await probe.main([])).toBe(0);
        expect(probe.spawnSync).toHaveBeenCalledExactlyOnceWith('powershell.exe', [
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', expect.any(String),
        ], { encoding: 'utf8', timeout: 15_000 });
        expect(probe.spawnSync.mock.calls[0][1][4]).toContain(`GetTypeFromProgID("${info.progId}")`);
        expect(probe.ran).toEqual(['safe', 'also-safe']);
    });

    it.each(hostModes)('uses one build process for $host, checkpoint=$checkpoint, and writes the marker before the registry read',
        async ({ host, info, checkpoint }) => {
            const probe = setup([{ id: 'crash', host, crashesHostBuilds: [build] },
                { id: 'also-crash', host, crashesHostBuilds: [build] }], checkpoint);
            probe.spawnSync.mockReturnValue({ status: 0, stdout: `${marker}\r\n${build}\r\n` });
            expect(await probe.main([])).toBe(0);
            expect(probe.spawnSync).toHaveBeenCalledExactlyOnceWith('powershell.exe', [
                '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', expect.any(String),
            ], { encoding: 'utf8', timeout: 15_000 });
            const script = probe.spawnSync.mock.calls[0][1][4];
            const lookup = script.indexOf(`$type = [type]::GetTypeFromProgID("${info.progId}")`);
            const nullCheck = script.indexOf('if ($null -eq $type) { exit 2 }');
            const write = script.indexOf(`[Console]::Out.WriteLine('${marker}')`);
            const flush = script.indexOf('[Console]::Out.Flush()');
            const registryRead = script.indexOf('[Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey($path)');
            const exeLine = script.indexOf("[Console]::Out.WriteLine('XLIDE-ORACLE-EXE ' + $exe)");
            const versionRead = script.indexOf('[Diagnostics.FileVersionInfo]::GetVersionInfo($exe)');
            expect(lookup).toBeGreaterThanOrEqual(0);
            expect(nullCheck).toBeGreaterThan(lookup);
            expect(write).toBeGreaterThan(nullCheck);
            expect(flush).toBeGreaterThan(write);
            expect(registryRead).toBeGreaterThan(flush);
            // The exe is reported before its version is read, so an unreadable
            // version still leaves the bitness readable.
            expect(exeLine).toBeGreaterThan(registryRead);
            expect(versionRead).toBeGreaterThan(exeLine);
            expect(probe.errors).toEqual([]);
            expect(probe.ran).toEqual([]);
            expect(JSON.parse(probe.output[0])).toMatchObject({ skippedCount: 2 });
        });

    it.each([
        { name: 'unregistered', status: 2 },
        { name: 'timeout', status: null, error: timeoutError },
        { name: 'ENOENT', status: null, error: new Error('spawnSync powershell.exe ENOENT') },
        { name: 'EMFILE', status: null, error: new Error('spawnSync powershell.exe EMFILE') },
    ])('keeps registration-only failure classification: $name', (completed) => {
        const probe = setup();
        probe.spawnSync.mockReturnValue(completed);
        expect(probe.comRegistered('Excel.Application')).toBe(false);
    });
});

const exePath = 'C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE';
const exeLine = `XLIDE-ORACLE-EXE ${exePath}`;

describe('VBE oracle host exe from the probe', () => {
    it.each([
        { name: 'exe then build, LF', completed: { status: 0, stdout: `${marker}\n${exeLine}\n${build}\n` },
            expected: { registered: true, build, exe: exePath } },
        { name: 'exe then build, CRLF', completed: { status: 0, stdout: `${marker}\r\n${exeLine}\r\n${build}\r\n` },
            expected: { registered: true, build, exe: exePath } },
        { name: 'exe then a version timeout', completed: { status: null, stdout: `${marker}\r\n${exeLine}\r\n`, error: timeoutError },
            expected: { registered: true, build: null, exe: exePath, why: timeoutError.message } },
        { name: 'exe then no version', completed: { status: 2, stdout: `${marker}\n${exeLine}\n` },
            expected: { registered: true, build: null, exe: exePath, why: 'PowerShell exited 2' } },
        { name: 'an empty exe line', completed: { status: 0, stdout: `${marker}\nXLIDE-ORACLE-EXE \n${build}\n` },
            expected: { registered: true, build } },
    ])('reads $name', ({ completed, expected }) => {
        const probe = setup();
        probe.spawnSync.mockReturnValueOnce(completed);
        expect(probe.registeredHostBuild('Excel.Application')).toEqual(expected);
    });

    it('hands the exe the probe found to the bitness check', async () => {
        const seen: unknown[] = [];
        const probe = setup([{ id: 'safe' }], false, (file: unknown) => { seen.push(file); return 64; });
        probe.spawnSync.mockReturnValue({ status: 0, stdout: `${marker}\r\n${exeLine}\r\n${build}\r\n` });
        expect(await probe.main([])).toBe(0);
        expect(seen).toEqual([exePath]);
        expect(probe.ran).toEqual(['safe']);
    });
});

describe('VBE oracle startup messages and exit codes', () => {
    it('rejects unknown hosts in case order before probing any known host', async () => {
        const probe = setup([{ id: 'excel' }, { id: 'unknown', host: 'other' },
            { id: 'also-unknown', host: 'another' }], true);
        expect(await probe.main([])).toBe(2);
        expect(probe.errors).toEqual(['Unknown oracle host: unknown names host "other"; also-unknown names host "another". Use excel, word, powerpoint, access, or leave host out for excel.']);
        expect(probe.spawnSync).not.toHaveBeenCalled();
    });

    const failures = classifications.filter((row) => row.build === null);
    it.each(hostModes.flatMap((mode) => failures.map((failure) => [
        `${mode.host}: ${failure.name}, checkpoint=${mode.checkpoint}`, { ...mode, ...failure },
    ] as const)))(
        '%s', async (_name, { host, info, checkpoint, completed, registered, why, fallback }) => {
            const probe = setup([{ id: 'crash', host, crashesHostBuilds: [build] }], checkpoint);
            probe.spawnSync.mockReturnValueOnce(completed).mockReturnValue({ status: fallback ? 0 : 2 });
            const fatal = !registered || checkpoint;
            expect(await probe.main([])).toBe(fatal ? 2 : 0);
            expect(probe.errors).toEqual(!registered ? [
                `VBE oracle cases for ${host} require Microsoft ${info.noun} COM registration.`,
            ] : checkpoint ? [
                `--checkpoint requires a readable file version for the registered ${info.processName}.EXE: ${why}`,
            ] : []);
            expect(probe.ran).toEqual(fatal ? [] : ['crash']);
            expect(probe.spawnSync).toHaveBeenCalledTimes(fallback === undefined ? 1 : 2);
        });

    // Every host gets the build probe now, for its exe's bitness.
    const paths = [
        { checkpoint: false, laterNeedsBuild: true },
        { checkpoint: true, laterNeedsBuild: true },
    ];
    const missing = [
        { name: 'exit 2', completed: { status: 2 } },
        { name: 'timeout before marker', completed: { status: null, error: timeoutError } },
    ];
    const earlierFailures = [
        { earlierName: 'after marker', earlierCompleted: { status: null, stdout: `${marker}\n`, error: timeoutError }, fallback: false },
        { earlierName: 'before marker with successful fallback', earlierCompleted: { status: null, stdout: null, error: timeoutError }, fallback: true },
    ];
    it.each(hosts.flatMap(([host, info]) => paths.flatMap((path) => missing.flatMap((failure) =>
        earlierFailures.map((earlier) => ({ host, info, ...path, ...failure, ...earlier }))))))(
        '$host registration wins over earlier build timeout $earlierName: $name, checkpoint=$checkpoint, laterNeedsBuild=$laterNeedsBuild',
        async ({ host, info, checkpoint, laterNeedsBuild, completed, earlierCompleted, fallback }) => {
            const earlierHost = hosts.find(([other]) => other !== host)![0];
            const probe = setup([{ id: 'earlier', host: earlierHost, crashesHostBuilds: [] },
                { id: 'missing', host, ...(laterNeedsBuild ? { crashesHostBuilds: [] } : {}) }], checkpoint);
            probe.spawnSync.mockReturnValue({ status: 2 }).mockReturnValueOnce(earlierCompleted);
            if (fallback) { probe.spawnSync.mockReturnValueOnce({ status: 0 }); }
            probe.spawnSync.mockReturnValueOnce(completed);
            expect(await probe.main([])).toBe(2);
            expect(probe.errors).toEqual([
                `VBE oracle cases for ${host} require Microsoft ${info.noun} COM registration.`,
            ]);
            expect(probe.ran).toEqual([]);
            expect(probe.spawnSync).toHaveBeenCalledTimes((laterNeedsBuild ? 3 : 2) + (fallback ? 1 : 0));
            expect(probe.spawnSync.mock.calls.map((call) => call[2])).toEqual([
                { encoding: 'utf8', timeout: 15_000 },
                ...(fallback ? [{ encoding: 'utf8', timeout: 5000 }] : []),
                { encoding: 'utf8', timeout: laterNeedsBuild ? 15_000 : 5000 },
                ...(laterNeedsBuild ? [{ encoding: 'utf8', timeout: 5000 }] : []),
            ]);
        });

    it('reports the first missing registration in case order and does not probe later hosts', async () => {
        const probe = setup([{ id: 'word', host: 'word' }, { id: 'excel' }], true);
        probe.spawnSync.mockReturnValue({ status: 2 });
        expect(await probe.main([])).toBe(2);
        expect(probe.errors).toEqual(['VBE oracle cases for word require Microsoft Word COM registration.']);
        expect(probe.spawnSync).toHaveBeenCalledTimes(2);
    });

    it('reports the first build error in case order after all registrations pass', async () => {
        const probe = setup([{ id: 'word', host: 'word' }, { id: 'excel' }], true);
        probe.spawnSync.mockImplementation(() => {
            expect(probe.errors).toEqual([]);
            return { status: 2, stdout: `${marker}\n` };
        });
        expect(await probe.main([])).toBe(2);
        expect(probe.errors).toEqual([
            '--checkpoint requires a readable file version for the registered WINWORD.EXE: PowerShell exited 2',
        ]);
        expect(probe.spawnSync).toHaveBeenCalledTimes(2);
    });
});
