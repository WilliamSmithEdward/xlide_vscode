import crypto from 'node:crypto';
import { EventEmitter } from 'node:events';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { runInNewContext } from 'node:vm';
import { describe, expect, it } from 'vitest';
import { OFFICE_HOST_APPS } from '../../src/officeHostApps';

// A fixture's `bitness`: which Office it was measured in. Run the
// coordinator's own functions, with no real PowerShell, Office or files.
const source = readFileSync(join(process.cwd(), 'syntax_corpus/oracle/run_vbe_oracle.mjs'), 'utf8');
const names = ['caseHost', 'unsupportedHost', 'malformedReferences', 'caseBitness', 'unsupportedBitness', 'exeBitness',
    'otherBitnessResult', 'hostsNeedingBuilds', 'knownHostCrashResult', 'expectedMatches', 'isOracleInfrastructureFailure',
    'promoteObservedCases', 'evidencePhaseForCase', 'evidencePhaseForResult', 'diagnosticMeaningForResult'];
const functions = names.map((name) => {
    const start = source.indexOf(`function ${name}(`);
    return source.slice(start, source.indexOf('\n}', start) + 2);
}).join('\n');
const officeHosts = Object.fromEntries(Object.entries(OFFICE_HOST_APPS).map(([host, info]) => [host,
    { progId: info.progId, image: info.processName, noun: info.noun }]));
const policy = runInNewContext(`const DEFAULT_HOST = 'excel';
const EVIDENCE_OUTCOMES = new Set(['accepted', 'rejected']);
${functions}\n({${names.join(',')}})`, { OFFICE_HOSTS: officeHosts, Buffer });
// The checkpoint hash itself, with a fixed harness hash.
const caseHash = runInNewContext(`const HARNESS_HASH = 'harness';
${source.slice(source.indexOf('function caseHash('), source.indexOf('\n}', source.indexOf('function caseHash(')) + 2)}
caseHash`, { crypto }) as (kase: unknown, build: string | null, bitness: number) => string;
const mainStart = source.indexOf('async function main(');
const mainSource = source.slice(mainStart, source.indexOf('\nprocess.exit(', mainStart));
const hosts = Object.keys(OFFICE_HOST_APPS);

/** The first bytes of a PE file whose COFF header names `machine`. */
function peHead(machine: number, { mz = 0x5a4d, signature = 0x00004550, offset = 0x80 } = {}): Buffer {
    const head = Buffer.alloc(0x200);
    head.writeUInt16LE(mz, 0);
    head.writeUInt32LE(offset, 0x3c);
    head.writeUInt32LE(signature, offset);
    head.writeUInt16LE(machine, offset + 4);
    return head;
}

type Case = { id: string; host?: string; bitness?: unknown; expected?: string; crashesHostBuilds?: string[] };

function setup(cases: Case[], { bitness = 64 as number | null | Record<string, number>, strict = true, json = true,
    checkpoint = undefined as string | undefined, exe = 'C:\\Office\\HOST.EXE' as string | null } = {}) {
    const trace: string[] = [];
    const output: string[] = [];
    const errors: string[] = [];
    const writes: string[] = [];
    const hashes: unknown[][] = [];
    const main = runInNewContext(`${mainSource}\nmain`, {
        ...policy,
        exeBitness: (file: string | undefined) => {
            trace.push(`exe:${file}`);
            if (!file) { return undefined; }
            const host = Object.keys(officeHosts).find((name) => file.includes(officeHosts[name].image));
            return bitness !== null && typeof bitness === 'object' ? bitness[host ?? 'excel'] : bitness ?? undefined;
        },
        parseArgs: () => ({ timeout: 30, timeoutRetries: 0, dialogHoldSeconds: 0, casesPerInstance: 40,
            caseIds: [], strict, json, checkpoint }),
        process: Object.assign(new EventEmitter(), { platform: 'win32' }),
        setDialogSounds: (action: string) => trace.push(`sounds:${action}`),
        console: { log: (line: string) => output.push(line), error: (line: string) => errors.push(line) },
        loadCasesDocument: () => ({ cases }),
        loadCheckpoint: () => new Map(),
        OFFICE_HOSTS: officeHosts,
        registeredHostBuild: (progId: string) => {
            trace.push(`probe:${progId}`);
            const host = Object.values(officeHosts).find((info) => info.progId === progId)!;
            return { registered: true, build: '16.0.20527.20088',
                ...(exe ? { exe: exe.replace('HOST', host.image) } : {}) };
        },
        acquireOracleRunLock: async () => { trace.push('lock'); return { held: true, release: async () => {} }; },
        oracleRunLockAddress: () => 'test-lock', oracleRunOwner: () => ({ pid: 1 }), describeRefusal: () => 'refused',
        OwnedHosts: class { load() { trace.push('owned:load'); return []; } add() {} sweep() { return []; } },
        ownedHostsPath: () => 'hosts.json', describeLeftHosts: () => 'left', abandonActiveCases: () => {},
        unwritableCheckpoint: () => undefined,
        OracleWorker: class {
            constructor() { trace.push('worker'); }
            async close() {}
            kill() {}
        },
        caseHash: (...args: unknown[]) => { hashes.push(args); return JSON.stringify(args.slice(1)); },
        checkpointKey: (id: string, hash: string) => `${id}\n${hash}`,
        phaseMatches: () => true,
        runCaseWithRetries: async (_worker: unknown, kase: Case) => {
            trace.push(`run:${kase.id}`);
            return { caseId: kase.id, outcome: 'accepted', stage: 'compile' };
        },
        fs: { appendFileSync: (_file: string, line: string) => writes.push(line), writeFileSync: () => {} },
    }) as (argv: string[]) => Promise<number>;
    return { main, trace, output, errors, writes, hashes };
}

describe('VBE oracle fixture bitness', () => {
    it('reads absent as 64 and accepts exactly 32 or 64', () => {
        expect(policy.caseBitness({ id: 'a' })).toBe(64);
        expect(policy.caseBitness({ id: 'a', bitness: 32 })).toBe(32);
        for (const bitness of [undefined, 32, 64]) {
            expect(policy.unsupportedBitness({ id: 'a', bitness })).toBeUndefined();
        }
        for (const bitness of [null, '32', 16, 128, 32.5, true, [32]]) {
            expect(policy.unsupportedBitness({ id: 'a', bitness }), JSON.stringify(bitness))
                .toBe(`a names bitness ${JSON.stringify(bitness)}`);
        }
    });

    it('reads 32 or 64 from the exe header, and nothing from anything else', () => {
        const read = (head: Buffer) => policy.exeBitness('C:\\Office\\EXCEL.EXE', () => head);
        expect(read(peHead(0x014c))).toBe(32);
        expect(read(peHead(0x8664))).toBe(64);
        expect(read(peHead(0x8664, { offset: 0x40 }))).toBe(64);
        expect(read(peHead(0xaa64))).toBeUndefined(); // ARM64
        expect(read(peHead(0x8664, { mz: 0 }))).toBeUndefined();
        expect(read(peHead(0x8664, { signature: 0 }))).toBeUndefined();
        expect(read(peHead(0x8664).subarray(0, 0x82))).toBeUndefined(); // header past the bytes read
        expect(read(Buffer.alloc(10))).toBeUndefined();
        expect(policy.exeBitness('C:\\missing.exe', () => { throw new Error('ENOENT'); })).toBeUndefined();
        expect(policy.exeBitness(undefined, () => peHead(0x8664))).toBeUndefined();
    });

    it.each(hosts)('skips a %s case for the other bitness with the reason', (host) => {
        const noun = officeHosts[host].noun;
        const map = new Map([[host, 64]]);
        expect(policy.otherBitnessResult({ id: 'a', host, bitness: 64 }, map)).toBeUndefined();
        expect(policy.otherBitnessResult({ id: 'a', host }, map)).toBeUndefined();
        expect(policy.otherBitnessResult({ id: 'a', host, bitness: 32 }, map)).toEqual({
            caseId: 'a', outcome: 'skipped', stage: 'other_bitness', hostBitness: 64, hresult: null,
            message: `Skipped: needs 32-bit Office; ${noun} here is 64-bit.`,
        });
        expect(policy.otherBitnessResult({ id: 'a', host }, new Map([[host, 32]])))
            .toMatchObject({ stage: 'other_bitness', message: `Skipped: needs 64-bit Office; ${noun} here is 32-bit.` });
        // A skip is no evidence: it cannot be promoted or fail an expectation.
        const skip = policy.otherBitnessResult({ id: 'a', host, bitness: 32 }, map);
        expect(policy.isOracleInfrastructureFailure(skip.outcome)).toBe(false);
        expect(policy.promoteObservedCases({ cases: [{ id: 'a', expected: 'observe' }] }, [skip]).errors)
            .toEqual(["a: outcome 'skipped' cannot be promoted"]);
    });
});

describe('VBE oracle checkpoint hash', () => {
    it('differs by the measured bitness of the same build', () => {
        const kase = { id: 'a', source: 'x' };
        expect(caseHash(kase, '16.0.20527.20088', 32)).not.toBe(caseHash(kase, '16.0.20527.20088', 64));
        expect(caseHash(kase, '16.0.20527.20088', 64)).toBe(caseHash({ ...kase }, '16.0.20527.20088', 64));
        expect(caseHash(kase, '16.0.20527.20088', 64)).not.toBe(caseHash(kase, '16.0.20527.20089', 64));
    });
});

describe('VBE oracle runs by bitness', () => {
    it('refuses a malformed bitness before any probe, the cases it would skip too', async () => {
        const probe = setup([{ id: 'fine' }, { id: 'bad', bitness: '32' }, { id: 'worse', bitness: 16 }]);
        expect(await probe.main([])).toBe(2);
        expect(probe.errors).toEqual(['Unknown oracle bitness: bad names bitness "32"; worse names bitness 16. Use 32 or 64, or leave bitness out for 64.']);
        expect(probe.trace).toEqual([]);
    });

    it.each([
        { name: 'no exe found', exe: null, why: 'its registered exe was not found' },
        { name: 'another machine type', exe: 'C:\\Office\\HOST.EXE', why: 'C:\\Office\\EXCEL.EXE is not an x86 or x64 program' },
    ])('refuses a run whose host bitness cannot be read: $name', async ({ exe, why }) => {
        const probe = setup([{ id: 'a' }], { bitness: null, exe });
        expect(await probe.main([])).toBe(2);
        expect(probe.errors).toEqual([`Cannot tell whether Excel here is 32-bit or 64-bit Office: ${why}.`]);
        expect(probe.trace.filter((step) => !step.startsWith('probe:') && !step.startsWith('exe:'))).toEqual([]);
    });

    it.each([32, 64])('runs the cases of a %s-bit machine and skips the others in order', async (machine) => {
        const other = machine === 32 ? 64 : 32;
        const probe = setup([{ id: 'same', bitness: machine, expected: 'accepted' }, { id: 'other', bitness: other, expected: 'accepted' },
            { id: 'absent', expected: 'accepted' }], { bitness: machine, checkpoint: 'checkpoint' });
        expect(await probe.main([])).toBe(0);
        const ran = probe.trace.filter((step) => step.startsWith('run:'));
        expect(ran).toEqual(machine === 64 ? ['run:same', 'run:absent'] : ['run:same']);
        const report = JSON.parse(probe.output[0]);
        expect(report.results.map((r: { caseId: string; outcome: string }) => `${r.caseId}:${r.outcome}`))
            .toEqual(machine === 64 ? ['same:accepted', 'other:skipped', 'absent:accepted'] : ['same:accepted', 'other:skipped', 'absent:skipped']);
        expect(report).toMatchObject({ failureCount: 0, skippedCount: machine === 64 ? 1 : 2, otherBitnessCount: machine === 64 ? 1 : 2 });
        // Each result says which Office answered it; skips are never checkpointed.
        for (const result of report.results) {
            expect(result.hostBitness).toBe(machine);
        }
        expect(probe.writes.map((line) => JSON.parse(line).result.caseId)).toEqual(machine === 64 ? ['same', 'absent'] : ['same']);
        expect(probe.writes.every((line) => JSON.parse(line).result.hostBitness === machine)).toBe(true);
        // The checkpoint hash covers the measured bitness with the build.
        expect(probe.hashes.every(([, build, bitness]) => build === '16.0.20527.20088' && bitness === machine)).toBe(true);
        expect(probe.trace).toContain('sounds:silence');
    });

    it.each([true, false])('starts nothing when every case is for the other bitness, json=%s', async (json) => {
        const probe = setup([{ id: 'a', bitness: 32, expected: 'accepted' }, { id: 'b', host: 'word', bitness: 32, expected: 'rejected' }],
            { json, strict: true });
        expect(await probe.main([])).toBe(0);
        expect(probe.trace.filter((step) => !step.startsWith('probe:') && !step.startsWith('exe:'))).toEqual([]);
        if (json) {
            expect(JSON.parse(probe.output[0])).toMatchObject({ failureCount: 0, skippedCount: 2, otherBitnessCount: 2 });
        } else {
            expect(probe.output.slice(0, 4)).toEqual(['SKIP a: outcome=skipped expected=accepted stage=other_bitness',
                '  Skipped: needs 32-bit Office; Excel here is 64-bit.',
                'SKIP b: outcome=skipped expected=rejected stage=other_bitness',
                '  Skipped: needs 32-bit Office; Word here is 64-bit.']);
            expect(probe.output.at(-1)).toBe('\n2 oracle case(s), 0 expectation failure(s), 0 oracle infrastructure failure(s), '
                + '2 skipped case(s) (2 for the other Office bitness; none ran).');
        }
    });

    it('judges each host by its own exe', async () => {
        const probe = setup([{ id: 'excel', bitness: 32 }, { id: 'word', host: 'word', bitness: 32 }],
            { bitness: { excel: 64, word: 32 } });
        expect(await probe.main([])).toBe(0);
        expect(probe.trace.filter((step) => step.startsWith('run:'))).toEqual(['run:word']);
    });

    it('counts a known host crash as a skip, not as the other bitness', async () => {
        const probe = setup([{ id: 'crash', crashesHostBuilds: ['16.0.20527.20088'] }, { id: 'other', bitness: 32 }]);
        expect(await probe.main([])).toBe(0);
        expect(JSON.parse(probe.output[0])).toMatchObject({ skippedCount: 2, otherBitnessCount: 1 });
        const results = JSON.parse(probe.output[0]).results as { stage: string; hostBitness: number }[];
        expect(results.map((r) => r.stage)).toEqual(['known_host_crash', 'other_bitness']);
        expect(results.map((r) => r.hostBitness)).toEqual([64, 64]);
    });
});
