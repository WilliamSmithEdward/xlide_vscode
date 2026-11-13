import { EventEmitter } from 'node:events';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { runInNewContext } from 'node:vm';
import { describe, expect, it } from 'vitest';

// Exercise the coordinator's pure policy without starting PowerShell or Office.
const source = readFileSync(join(process.cwd(), 'syntax_corpus/oracle/run_vbe_oracle.mjs'), 'utf8');
const names = ['caseHost', 'unsupportedHost', 'malformedReferences', 'validateCrashBuilds', 'hostsNeedingBuilds', 'knownHostCrashResult',
    'expectedMatches', 'isOracleInfrastructureFailure', 'isObserveWithoutVerdict', 'promoteObservedCases', 'formatCasesDocument',
    'evidencePhaseForCase', 'evidencePhaseForResult', 'diagnosticMeaningForResult', 'phaseMatches'];
const functions = names.map((name) => {
    const start = source.indexOf(`function ${name}(`);
    return source.slice(start, source.indexOf('\n}', start) + 2);
}).join('\n');
const policy = runInNewContext(`const DEFAULT_HOST = 'excel';
const OFFICE_HOSTS = { excel: {}, word: {}, powerpoint: {}, access: {} };
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
    it.each(['excel', 'word', 'powerpoint', 'access'])('promotes %s through the coordinator writer without changing other fixtures', async (host) => {
        for (const outcome of ['accepted', 'rejected', 'worker_error', 'timeout', 'skipped']) {
            const target = { id: 'target', host, mode: 'compile', expected: 'observe',
                source: 'n = 1\u00A0+ 1 \u00C9tat \uD83D\uDE00\u2028\u007F' };
            const untouched = { id: 'untouched', expected: 'observe', source: '\u00E9',
                provenance: 'observed-not-asserted' };
            const document = { cases: [target, untouched] };
            const before = JSON.stringify(document);
            const writes: Array<{ file: string; text: string; encoding: string }> = [];
            const ran: string[] = [];
            const start = source.indexOf('async function main(');
            const mainSource = source.slice(start, source.indexOf('\nprocess.exit(', start));
            const main = runInNewContext(`${mainSource}\nmain`, {
                ...policy,
                parseArgs: () => ({ timeout: 30, timeoutRetries: 0, dialogHoldSeconds: 0,
                    casesPerInstance: 40, caseIds: ['target'], cases: 'memory.json',
                    promoteObserved: true, strict: true, json: true }),
                process: Object.assign(new EventEmitter(), { platform: 'win32' }),
                console: { log() {}, error() {} },
                loadCasesDocument: () => document,
                OFFICE_HOSTS: { [host]: { progId: `${host}.Application` } },
                comRegistered: () => true,
                // No real Office, registry, or filesystem effects in this test.
                setDialogSounds() {},
                // The run lock is covered in vbeOracleRunLock.test.ts; here it is always free.
                acquireOracleRunLock: async () => ({ held: true, release: async () => {} }),
                oracleRunLockAddress: () => 'test-lock', oracleRunOwner: () => ({ pid: 1 }), describeRefusal: () => 'refused',
                // Owned hosts are covered in vbeOracleOwnedHosts.test.ts; here the run owns none.
                OwnedHosts: class { load() { return []; } add() {} sweep() { return []; } },
                ownedHostsPath: () => 'hosts.json', describeLeftHosts: () => 'left', abandonActiveCases: () => {},
                unwritableCheckpoint: () => undefined, // A real path is checked in vbeOracleOptions.test.ts.
                OracleWorker: class { async close() {} },
                caseHash: () => 'hash',
                runCaseWithRetries: async (_worker: unknown, kase: { id: string }) => {
                    ran.push(kase.id);
                    return { caseId: kase.id, outcome, stage: 'compile' };
                },
                fs: { writeFileSync: (file: string, text: string, encoding: string) => {
                    writes.push({ file, text, encoding });
                } },
            });

            const exit = await main([]);
            expect(ran).toEqual(['target']);
            if (outcome !== 'accepted' && outcome !== 'rejected') {
                expect(exit).toBe(1);
                expect(writes).toEqual([]);
                expect(JSON.stringify(document)).toBe(before);
                continue;
            }
            expect(exit).toBe(0);
            expect(writes).toHaveLength(1);
            const { file, text, encoding } = writes[0];
            expect(file).toBe('memory.json');
            expect(encoding).toBe('utf8');
            expect(/^[\n -~]*$/.test(text)).toBe(true);
            expect(text).toContain('\\u00A0');
            expect(text).toContain('\\uD83D\\uDE00');
            const written = JSON.parse(text);
            expect(written.cases[0]).toEqual({ ...JSON.parse(before).cases[0], expected: outcome,
                provenance: 'vbe-oracle-verified', evidencePhase: 'compile',
                diagnosticMeaning: outcome === 'accepted' ? 'compile-valid' : 'compile-error' });
            expect(written.cases[1]).toEqual(JSON.parse(before).cases[1]);
            expect(text.endsWith('\n')).toBe(true);
        }
    });

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
        // The run lock is covered in vbeOracleRunLock.test.ts; here it is always free.
        acquireOracleRunLock: async () => ({ held: true, release: async () => {} }),
        oracleRunLockAddress: () => 'test-lock', oracleRunOwner: () => ({ pid: 1 }), describeRefusal: () => 'refused',
        // Owned hosts are covered in vbeOracleOwnedHosts.test.ts; here the run owns none.
        OwnedHosts: class { load() { return []; } add() {} sweep() { return []; } },
        ownedHostsPath: () => 'hosts.json', describeLeftHosts: () => 'left', abandonActiveCases: () => {},
        unwritableCheckpoint: () => undefined, // A real path is checked in vbeOracleOptions.test.ts.
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
        // The run lock is covered in vbeOracleRunLock.test.ts; here it is always free.
        acquireOracleRunLock: async () => ({ held: true, release: async () => {} }),
        oracleRunLockAddress: () => 'test-lock', oracleRunOwner: () => ({ pid: 1 }), describeRefusal: () => 'refused',
        // Owned hosts are covered in vbeOracleOwnedHosts.test.ts; here the run owns none.
        OwnedHosts: class { load() { return []; } add() {} sweep() { return []; } },
        ownedHostsPath: () => 'hosts.json', describeLeftHosts: () => 'left', abandonActiveCases: () => {},
        unwritableCheckpoint: () => undefined, // A real path is checked in vbeOracleOptions.test.ts.
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


describe('VBE oracle invocation evidence', () => {
    it.each([
        ['excel', -2146233087, 'Exception calling "Run": "Value does not fall within the expected range."'],
        ['word', -2147024809, 'Value does not fall within the expected range.'],
        ['powerpoint', -2147023170, 'The remote procedure call failed.'],
        ['access', -2147024809, 'Value does not fall within the expected range.'],
    ])('does not promote a failed %s invocation or partially promote its batch', (host, hresult, message) => {
        const document = { cases: [
            { id: 'control', expected: 'observe', provenance: 'observed-not-asserted' },
            { id: 'invocation', expected: 'observe', provenance: 'observed-not-asserted' },
        ] };
        const before = JSON.stringify(document);
        const failed = { caseId: 'invocation', host, hostBuild: build, outcome: 'worker_error', stage: 'run', hresult, message };
        expect(policy.isOracleInfrastructureFailure(failed.outcome)).toBe(true);
        expect(policy.expectedMatches('accepted', failed.outcome)).toBe(false);
        expect(policy.expectedMatches('observe', failed.outcome)).toBe(false);
        expect(policy.promoteObservedCases(document, [
            { caseId: 'control', outcome: 'accepted', stage: 'run' }, failed,
        ])).toEqual({ promoted: 0, errors: ["invocation: outcome 'worker_error' cannot be promoted"] });
        expect(JSON.stringify(document)).toBe(before);
    });

    it.each(['issue609_13_runtime', 'issue609_17_runtime'])('keeps %s as an observation after the invocation recheck', id => {
        const corpus = JSON.parse(readFileSync(join(process.cwd(), 'syntax_corpus/oracle/vbe_oracle_cases.json'), 'utf8'));
        const fixture = corpus.cases.find((kase: { id: string }) => kase.id === id);
        expect(fixture).toMatchObject({ expected: 'observe', provenance: 'observed-not-asserted', evidencePhase: 'runtime', diagnosticMeaning: 'observation' });
        const before = JSON.stringify(fixture);
        expect(policy.promoteObservedCases({ cases: [fixture] }, [{ caseId: id, outcome: 'worker_error', stage: 'run' }]).promoted).toBe(0);
        expect(JSON.stringify(fixture)).toBe(before);
    });
});

describe('VBE oracle observe-only cases without a verdict', () => {
    const corpus = JSON.parse(readFileSync(join(process.cwd(), 'syntax_corpus/oracle/vbe_oracle_cases.json'), 'utf8'));
    const fixture = (id: string) => corpus.cases.find((kase: { id: string }) => kase.id === id);
    const hosts = ['excel', 'word', 'powerpoint', 'access'];
    // A build per host, so that a lookup in the wrong host's build shows.
    const builds: Record<string, string> = { excel: '16.0.20527.20196', word: '16.0.20527.20197',
        powerpoint: '16.0.20527.20198', access: '16.0.20527.20199' };
    const invalidArg = { outcome: 'worker_error', stage: 'run', hresult: -2146233087,
        message: 'Exception calling "Run" with "1" argument(s): "Value does not fall within the expected range."' };
    // The coordinator's own checkpoint reader, over the lines this run appends.
    const readCheckpoint = (lines: string[]) => runInNewContext(`const EVIDENCE_OUTCOMES = new Set(['accepted', 'rejected']);
${['checkpointKey', 'loadCheckpoint'].map((name) => {
        const start = source.indexOf(`function ${name}(`);
        return source.slice(start, source.indexOf('\n}', start) + 2);
    }).join('\n')}\nloadCheckpoint('checkpoint')`, {
        fs: { readFileSync: () => lines.join(''), appendFileSync() {} },
    });

    type Case = { id: string; host?: string; expected?: string };
    function runMain(cases: Case[], replies: Record<string, object>,
        options: { json?: boolean; checkpoint?: string; promoteObserved?: boolean; caseIds?: string[] } = {}) {
        const output: string[] = [];
        const errors: string[] = [];
        const ran: string[] = [];
        const appended: string[] = [];
        const writes: string[] = [];
        const start = source.indexOf('async function main(');
        const mainSource = source.slice(start, source.indexOf('\nprocess.exit(', start));
        const main = runInNewContext(`${mainSource}\nmain`, {
            ...policy,
            parseArgs: () => ({ timeout: 30, timeoutRetries: 0, dialogHoldSeconds: 0, casesPerInstance: 40,
                caseIds: options.caseIds ?? [], cases: 'memory.json', checkpoint: options.checkpoint ?? null,
                strict: true, json: options.json ?? true, promoteObserved: options.promoteObserved ?? false }),
            process: Object.assign(new EventEmitter(), { platform: 'win32' }),
            setDialogSounds: () => {}, // Sound lifecycle is covered in vbeOracleDialogSounds.test.ts.
            console: { log: (line: string) => output.push(line), error: (line: string) => errors.push(line) },
            loadCasesDocument: () => ({ cases }),
            OFFICE_HOSTS: Object.fromEntries(hosts.map((host) => [host, { progId: `${host}.Application` }])),
            comRegistered: () => true,
            registeredHostBuild: (progId: string) => ({ registered: true, build: builds[progId.split('.')[0]] }),
            loadCheckpoint: () => readCheckpoint(appended),
            // The run lock is covered in vbeOracleRunLock.test.ts; here it is always free.
            acquireOracleRunLock: async () => ({ held: true, release: async () => {} }),
            oracleRunLockAddress: () => 'test-lock', oracleRunOwner: () => ({ pid: 1 }), describeRefusal: () => 'refused',
            // Owned hosts are covered in vbeOracleOwnedHosts.test.ts; here the run owns none.
            OwnedHosts: class { load() { return []; } add() {} sweep() { return []; } },
            ownedHostsPath: () => 'hosts.json', describeLeftHosts: () => 'left', abandonActiveCases: () => {},
            unwritableCheckpoint: () => undefined, // A real path is checked in vbeOracleOptions.test.ts.
            OracleWorker: class { async close() {} },
            caseHash: () => 'hash', checkpointKey: (id: string, hash: string) => `${id}\n${hash}`,
            runCaseWithRetries: async (_worker: unknown, kase: Case) => {
                ran.push(kase.id);
                return { caseId: kase.id, hostBuild: builds[kase.host ?? 'excel'],
                    ...(replies[kase.id] ?? { outcome: 'accepted', stage: 'compile' }) };
            },
            fs: {
                appendFileSync: (_file: string, line: string) => appended.push(line),
                writeFileSync: (_file: string, text: string) => writes.push(text),
            },
        });
        return { main, output, errors, ran, appended, writes };
    }

    // The two Excel fixtures, or the same source in another host, then an
    // ordinary compile case and run case.
    const issueCases = (host = 'excel') => [
        ...['issue609_13_runtime', 'issue609_17_runtime'].map((id) => (host === 'excel' ? fixture(id) : { ...fixture(id), host })),
        { id: 'later_compile', ...(host === 'excel' ? {} : { host }), mode: 'compile', expected: 'accepted' },
        { id: 'later_run', ...(host === 'excel' ? {} : { host }), mode: 'run', entryPoint: 'Main', expected: 'rejected' }];
    const issueReplies = { issue609_13_runtime: invalidArg, issue609_17_runtime: invalidArg,
        later_run: { outcome: 'rejected', stage: 'runtime_dialog' } };
    const issueIds = ['issue609_13_runtime', 'issue609_17_runtime', 'later_compile', 'later_run'];

    it.each(hosts.flatMap((host) => [[host, false], [host, true]]))('reports both invocation failures in %s and runs the cases after them, checkpoint=%s', async (host, checkpoint) => {
        const run = runMain(issueCases(host as string), issueReplies, checkpoint ? { checkpoint: 'checkpoint' } : {});
        expect(await run.main([])).toBe(0);
        expect(run.ran).toEqual(issueIds);
        const report = JSON.parse(run.output[0]);
        expect(report).toMatchObject({ failureCount: 0, oracleFailureCount: 0, skippedCount: 0, noVerdictCount: 2 });
        for (const result of report.results.slice(0, 2)) {
            // Still a worker error, never a verdict (#1402).
            expect(result).toMatchObject({ outcome: 'worker_error', stage: 'run', expected: 'observe', matched: false, noVerdict: true });
        }
        expect(report.results.slice(2).map((result: { matched: boolean; noVerdict?: boolean }) => [result.matched, result.noVerdict]))
            .toEqual([[true, undefined], [true, undefined]]);
    });

    it('marks them NO-VERDICT in text and counts them apart from failures', async () => {
        const run = runMain(issueCases(), issueReplies, { json: false });
        expect(await run.main([])).toBe(0);
        const text = run.output.join('\n');
        expect(text).toContain('NO-VERDICT issue609_13_runtime: outcome=worker_error expected=observe stage=run');
        expect(text).toContain('NO-VERDICT issue609_17_runtime: outcome=worker_error expected=observe stage=run');
        expect(text).toContain('Value does not fall within the expected range.');
        expect(text).toContain('PASS later_run: outcome=rejected');
        expect(text).not.toContain('ORACLE-FAIL');
        expect(text).toContain('4 oracle case(s), 0 expectation failure(s), 0 oracle infrastructure failure(s), '
            + '0 skipped case(s), 2 observe-only case(s) without a verdict.');
    });

    // Run's answer cannot be told from these by the coordinator; all are reported.
    it.each([
        ['a host that crashed during the run', 'Excel (pid 4242) exited during the run stage. A crash is not VBA evidence.'],
        ['a macro Run cannot find', 'Exception calling "Run" with "1" argument(s): "Cannot run the macro \'Main\'."'],
    ])('reports %s in an observe case and goes on', async (_why, message) => {
        const run = runMain(issueCases(), { ...issueReplies, issue609_13_runtime: { outcome: 'worker_error', stage: 'run', message } });
        expect(await run.main([])).toBe(0);
        expect(run.ran).toEqual(issueIds);
        expect(JSON.parse(run.output[0]).results[0]).toMatchObject({ noVerdict: true, message });
    });

    it.each(['accepted', 'rejected'])('still stops at a worker error in a case expected %s', async (expected) => {
        const run = runMain([{ ...fixture('issue609_13_runtime'), expected }, { id: 'after', expected: 'accepted' }],
            { issue609_13_runtime: invalidArg }, { json: false });
        expect(await run.main([])).toBe(1);
        expect(run.ran).toEqual(['issue609_13_runtime']);
        expect(run.output.join('\n')).toContain(`ORACLE-FAIL issue609_13_runtime: outcome=worker_error expected=${expected}`);
        expect(run.output.at(-1)).toContain('1 oracle infrastructure failure(s), 0 skipped case(s), 0 observe-only case(s) without a verdict.');
    });

    const stops = (kase: Case, reply: object, options: { checkpoint?: string } = {}) => async () => {
        const run = runMain([kase, { id: 'after', host: kase.host, expected: 'accepted' }], { [kase.id]: reply }, options);
        expect(await run.main([])).toBe(1);
        expect(run.ran).toEqual([kase.id]);
        expect(JSON.parse(run.output[0])).toMatchObject({ oracleFailureCount: 1, noVerdictCount: 0 });
    };

    it.each([
        ['timed out every attempt', { outcome: 'oracle_failure', stage: 'run', message: 'Oracle timed out after 3 attempt(s) at 45 seconds each.' }],
        ['could not start its host', { outcome: 'worker_error', stage: 'start_host', hostBuild: null, message: 'Could not tell which Excel process the oracle started.' }],
        ['could not create its file', { outcome: 'worker_error', stage: 'create_file', message: 'Exception from HRESULT: 0x800A03EC' }],
        ['failed before its host', { outcome: 'worker_error', stage: 'setup', hostBuild: null, message: 'Unexpected character.' }],
        ['lost its worker', { outcome: 'worker_error', stage: 'worker', hostBuild: undefined, message: 'PowerShell exited 1' }],
        ['found no Compile command', { outcome: 'worker_error', stage: 'compile', message: "Could not find exact VBE command 'Compile VBAProject'. Oracle compile command was not invoked." }],
        ['met a Security Notice', { outcome: 'worker_error', stage: 'security_notice', message: "Office showed 'Microsoft Excel Security Notice' while adding the project reference." }],
        ['could not reach its project', { outcome: 'worker_error', stage: 'add_module', message: 'Programmatic access to Visual Basic Project is not trusted.' }],
        ['met a host left running', { outcome: 'worker_error', stage: 'host_cleanup', hostBuild: undefined, message: 'EXCEL.EXE pid 4242 is still running.' }],
    ])('still stops at an observe case that %s', (_why, reply) => stops(fixture('issue609_13_runtime'), reply)());

    it('still stops at an observe run case with no entry point', () => {
        const { entryPoint: _none, ...kase } = fixture('issue609_13_runtime');
        return stops(kase, { outcome: 'worker_error', stage: 'run', message: 'A run case needs an entryPoint.' })();
    });

    it.each(hosts)('still stops at an observe %s case that reports another host\'s build or none in a checkpoint run', async (host) => {
        const kase = { ...fixture('issue609_13_runtime'), host };
        const other = hosts.find((one) => one !== host) as string;
        await stops(kase, { ...invalidArg, hostBuild: builds[other] }, { checkpoint: 'checkpoint' })();
        await stops(kase, { ...invalidArg, hostBuild: null }, { checkpoint: 'checkpoint' })();
    });

    it('takes the build check from the coordinator\'s own conversion', async () => {
        const start = source.indexOf('async function runCaseOnce(');
        const runCaseOnceSource = source.slice(start, source.indexOf('\n}', start) + 2);
        const kase = fixture('issue609_13_runtime');
        const expected = builds.excel;
        const convert = async (reply: object) => runInNewContext(`${runCaseOnceSource}\nrunCaseOnce`, {
            ...policy, OFFICE_HOSTS: { excel: { noun: 'Excel' } }, EVIDENCE_OUTCOMES: new Set(['accepted', 'rejected']),
            runCaseOnceUnchecked: async () => ({ caseId: kase.id, ...reply }),
        })({ hostBuilds: new Map([['excel', expected]]) }, kase, 45, 0);
        // Evidence from another build or none becomes a worker error that stops the run.
        for (const hostBuild of [builds.word, undefined]) {
            const result = await convert({ outcome: 'accepted', stage: 'run', hostBuild });
            expect(result.outcome).toBe('worker_error');
            expect(policy.isObserveWithoutVerdict(kase, result, expected)).toBe(false);
        }
        const refused = await convert({ ...invalidArg, hostBuild: expected });
        expect(refused).toMatchObject(invalidArg);
        expect(policy.isObserveWithoutVerdict(kase, refused, expected)).toBe(true);
        expect(policy.isObserveWithoutVerdict({ ...kase, expected: 'rejected' }, refused, expected)).toBe(false);
    });

    it('never reuses the invocation failure from a checkpoint, so a resumed run tries it again', async () => {
        const first = runMain(issueCases(), issueReplies, { checkpoint: 'checkpoint' });
        expect(await first.main([])).toBe(0);
        expect(first.appended.map((line) => JSON.parse(line).result.outcome))
            .toEqual(['worker_error', 'worker_error', 'accepted', 'rejected']);
        expect([...readCheckpoint(first.appended).keys()]).toEqual(['later_compile\nhash', 'later_run\nhash']);

        const resumed = runMain(issueCases(), issueReplies, { checkpoint: 'checkpoint' });
        resumed.appended.push(...first.appended);
        expect(await resumed.main([])).toBe(0);
        expect(resumed.ran).toEqual(['issue609_13_runtime', 'issue609_17_runtime']);
        expect(JSON.parse(resumed.output[0])).toMatchObject({ noVerdictCount: 2 });
    });

    it('refuses to promote it and writes nothing', async () => {
        const cases = issueCases();
        const before = JSON.stringify(cases);
        const run = runMain(cases, issueReplies, { promoteObserved: true, caseIds: ['issue609_13_runtime'] });
        expect(await run.main([])).toBe(1);
        expect(run.errors).toEqual(['Cannot promote observed oracle result(s):',
            "  issue609_13_runtime: outcome 'worker_error' cannot be promoted"]);
        expect(run.writes).toEqual([]);
        expect(JSON.stringify(cases)).toBe(before);
    });
});
