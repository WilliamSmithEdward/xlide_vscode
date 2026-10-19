import { EventEmitter } from 'node:events';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { runInNewContext } from 'node:vm';
import { describe, expect, it } from 'vitest';

// Exercise the coordinator's pure policy without starting PowerShell or Office.
const source = readFileSync(join(process.cwd(), 'syntax_corpus/oracle/run_vbe_oracle.mjs'), 'utf8');
const names = ['caseHost', 'validateCrashBuilds', 'hostsNeedingBuilds', 'knownHostCrashResult',
    'expectedMatches', 'isOracleInfrastructureFailure', 'promoteObservedCases', 'formatCasesDocument'];
const functions = names.map((name) => {
    const start = source.indexOf(`function ${name}(`);
    return source.slice(start, source.indexOf('\n}', start) + 2);
}).join('\n');
const policy = runInNewContext(`const DEFAULT_HOST = 'excel';
const EVIDENCE_OUTCOMES = new Set(['accepted', 'rejected']);
${functions}\n({${names.join(',')}})`);
const build = '16.0.20430.20182';

describe('VBE oracle known host crashes', () => {
    it.each(['excel', 'word', 'powerpoint', 'access'])('skips only an exact registered %s build', (host) => {
        const kase = { id: 'crash', host, crashesHostBuilds: [build] };
        const result = policy.knownHostCrashResult(kase, new Map([[host, build]]));
        expect(result).toMatchObject({ caseId: 'crash', outcome: 'skipped', stage: 'known_host_crash', hostBuild: build });
        expect(result.message).toContain(build);
        expect(policy.isOracleInfrastructureFailure(result.outcome)).toBe(false);
        expect(policy.expectedMatches('accepted', result.outcome)).toBe(false);
        expect(policy.knownHostCrashResult(kase, new Map([[host, '16.0.20430.20183']]))).toBeUndefined();
        expect(policy.knownHostCrashResult(kase, new Map())).toBeUndefined();
    });

    it('reads only affected hosts without a checkpoint and all hosts with one', () => {
        const cases = [{ id: 'a', crashesHostBuilds: [build] }, { id: 'b', host: 'word' },
            { id: 'c', host: 'access', crashesHostBuilds: [] }];
        expect(policy.hostsNeedingBuilds(cases, false)).toEqual(['excel', 'access']);
        expect(policy.hostsNeedingBuilds(cases, true)).toEqual(['excel', 'word', 'access']);
        expect(policy.knownHostCrashResult(cases[1], new Map([['word', build]]))).toBeUndefined();
    });

    it.each([null, build, [null], [12], [''], ['16.0.20430'], [' 16.0.20430.20182'], ['16.0.20430.20182\n']])('rejects malformed build lists: %j', (value) => {
        expect(() => policy.validateCrashBuilds([{ id: 'bad', crashesHostBuilds: value }])).toThrow(/bad: crashesHostBuilds/);
    });

    it('accepts absent and empty lists and refuses to promote skips', () => {
        expect(() => policy.validateCrashBuilds([{ id: 'a' }, { id: 'b', crashesHostBuilds: [] },
            { id: 'c', crashesHostBuilds: [build] }])).not.toThrow();
        const document = { cases: [{ id: 'a', expected: 'observe' }] };
        expect(policy.promoteObservedCases(document, [{ caseId: 'a', outcome: 'skipped' }]).errors)
            .toEqual(["a: outcome 'skipped' cannot be promoted"]);
        expect(document.cases[0].expected).toBe('observe');
        expect(policy.isOracleInfrastructureFailure('worker_error')).toBe(true);
    });
});

describe('VBE oracle promotion writes', () => {
    it('rewrites the committed fixture file byte for byte', () => {
        // \r appears on autocrlf checkouts, never in the blob.
        const text = readFileSync(join(process.cwd(), 'syntax_corpus/oracle/vbe_oracle_cases.json'), 'utf8')
            .replace(/\r\n/g, '\n');
        expect(policy.formatCasesDocument(JSON.parse(text)) === text,
            'lay the file out as JSON.stringify(document, null, 2) with DEL and non-ASCII written as \\uXXXX').toBe(true);
    });

    it('keeps non-ASCII characters escaped', () => {
        const source = 'n = 1\u00A0+ 1 \u00C9tat \uD83D\uDE00\u2028\u007F';
        const document = { cases: [{ id: 'a', expected: 'accepted', source }] };
        const text = policy.formatCasesDocument(document);
        expect(text).toContain('"source": "n = 1\\u00A0+ 1 \\u00C9tat \\uD83D\\uDE00\\u2028\\u007F"');
        expect(/^[\n -~]*$/.test(text)).toBe(true);
        expect(JSON.parse(text).cases[0].source).toBe(source);
    });
});

function setupUncheckpointedMain(registered: { build: string | null; why?: string }, outcome = 'accepted') {
    const output: string[] = [];
    const ran: string[] = [];
    const cases = [{ id: 'crash', crashesHostBuilds: [build], expected: 'accepted' },
        { id: 'safe', expected: 'accepted' }];
    const start = source.indexOf('async function main(');
    const mainSource = source.slice(start, source.indexOf('\nprocess.exit(', start));
    const main = runInNewContext(`${mainSource}\nmain`, {
        ...policy,
        parseArgs: () => ({ timeout: 30, timeoutRetries: 0, dialogHoldSeconds: 0,
            casesPerInstance: 40, caseIds: [], strict: true, json: true }),
        process: Object.assign(new EventEmitter(), { platform: 'win32' }),
        setDialogSounds: () => {}, // Sound lifecycle is covered in vbeOracleDialogSounds.test.ts.
        console: { log: (line: string) => output.push(line), error: (line: string) => output.push(line) },
        loadCasesDocument: () => ({ cases }),
        OFFICE_HOSTS: { excel: { progId: 'Excel.Application' } },
        comRegistered: () => true,
        registeredHostBuild: () => ({ registered: true, ...registered }),
        OracleWorker: class { async close() {} },
        caseHash: () => 'hash',
        runCaseWithRetries: async (_worker: unknown, kase: { id: string }) => {
            ran.push(kase.id);
            return { caseId: kase.id, outcome: kase.id === 'crash' ? outcome : 'accepted' };
        },
        phaseMatches: () => true,
    });
    return { main, output, ran };
}

it('runs known-crash fixtures normally when the registered build is unreadable without a checkpoint', async () => {
    const { main, output, ran } = setupUncheckpointedMain({ build: null, why: 'File version unavailable' });
    expect(await main([])).toBe(0);
    expect(ran).toEqual(['crash', 'safe']);
    expect(JSON.parse(output[0])).toMatchObject({ skippedCount: 0 });
});

it('stops on a worker crash when the registered build is not listed as a known crash', async () => {
    const { main, output, ran } = setupUncheckpointedMain({ build: '16.0.20430.20183' }, 'worker_error');
    expect(await main([])).toBe(1);
    expect(ran).toEqual(['crash']);
    expect(JSON.parse(output[0])).toMatchObject({ oracleFailureCount: 1, skippedCount: 0 });
});

it.each([true, false])('continues after a skip and excludes it from checkpoints and strict failures, json=%s', async (json) => {
    const writes: string[] = [];
    const output: string[] = [];
    const ran: string[] = [];
    const cases = [{ id: 'crash', crashesHostBuilds: [build], expected: 'accepted' },
        { id: 'safe', expected: 'accepted' }];
    const start = source.indexOf('async function main(');
    const mainSource = source.slice(start, source.indexOf('\nprocess.exit(', start));
    const main = runInNewContext(`${mainSource}\nmain`, {
        ...policy,
        parseArgs: () => ({ timeout: 30, timeoutRetries: 0, dialogHoldSeconds: 0,
            casesPerInstance: 40, caseIds: [], checkpoint: 'checkpoint', strict: true, json }),
        process: Object.assign(new EventEmitter(), { platform: 'win32' }),
        setDialogSounds: () => {}, // Sound lifecycle is covered in vbeOracleDialogSounds.test.ts.
        console: { log: (line: string) => output.push(line), error: (line: string) => output.push(line) },
        loadCasesDocument: () => ({ cases }),
        OFFICE_HOSTS: { excel: { progId: 'Excel.Application' } },
        comRegistered: () => true,
        registeredHostBuild: () => ({ registered: true, build }),
        loadCheckpoint: () => new Map(),
        OracleWorker: class { async close() {} },
        caseHash: () => 'hash', checkpointKey: () => 'key',
        runCaseWithRetries: async (_worker: unknown, kase: { id: string }) => {
            ran.push(kase.id);
            return { caseId: kase.id, outcome: 'accepted' };
        },
        phaseMatches: () => true,
        fs: { appendFileSync: (_file: string, line: string) => writes.push(line) },
    });
    expect(await main([])).toBe(0);
    expect(ran).toEqual(['safe']);
    expect(writes).toHaveLength(1);
    expect(JSON.parse(writes[0]).result.caseId).toBe('safe');
    if (json) {
        expect(JSON.parse(output[0])).toMatchObject({ failureCount: 0, oracleFailureCount: 0, skippedCount: 1 });
    } else {
        expect(output.join('\n')).toContain('SKIP crash');
        expect(output.join('\n')).toContain('1 skipped case(s)');
    }
});
