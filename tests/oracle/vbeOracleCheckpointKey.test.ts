import crypto from 'node:crypto';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { runInNewContext } from 'node:vm';
import { afterEach, describe, expect, it } from 'vitest';

// The checkpoint key: what it hashes, what it leaves out, and how a
// checkpoint written under the earlier key is treated.
const source = readFileSync(join(process.cwd(), 'syntax_corpus/oracle/run_vbe_oracle.mjs'), 'utf8');
const worker = readFileSync(join(process.cwd(), 'syntax_corpus/oracle/vbe_oracle_worker.ps1'), 'utf8');
const cases: Record<string, unknown>[] = JSON.parse(readFileSync(
    join(process.cwd(), 'syntax_corpus/oracle/vbe_oracle_cases.json'), 'utf8').replace(/^\uFEFF/, '')).cases;

// A top-level function, up to the next top-level declaration, so a closing
// brace at the start of a line inside a template literal cannot cut it short.
function slice(name: string) {
    const start = source.search(new RegExp(`^(?:async )?function ${name}\\(`, 'm'));
    expect(start).toBeGreaterThanOrEqual(0);
    const next = source.slice(start + 1).search(/\n(?:async function |function |const |let |class |\/\/ )/);
    return source.slice(start, start + 1 + next);
}
// A top-level constant, up to the next blank line, in either line ending.
function constant(name: string) {
    const start = source.search(new RegExp(`^const ${name} =`, 'm'));
    expect(start).toBeGreaterThanOrEqual(0);
    const end = source.slice(start).search(/\r?\n\r?\n/);
    expect(end).toBeGreaterThan(0);
    return source.slice(start, start + end);
}

const names = ['caseHash', 'checkpointKey', 'loadCheckpoint',
    'promoteObservedCases', 'evidencePhaseForCase', 'evidencePhaseForResult', 'diagnosticMeaningForResult'];
const constants = ['CHECKPOINT_KEY_VERSION', 'UNHASHED_FIELDS'];

function load() {
    const errors: string[] = [];
    const api = runInNewContext(`
        const EVIDENCE_OUTCOMES = new Set(['accepted', 'rejected']);
        const HARNESS_HASH = 'harness';
        ${constants.map(constant).join('\n')}
        ${names.map(slice).join('\n')}
        ({ ${[...names, ...constants].join(', ')} })`, {
        crypto,
        fs: { readFileSync, appendFileSync: (file: string, text: string) => writeFileSync(file, readFileSync(file, 'utf8') + text) },
        console: { error: (line: string) => errors.push(line) },
    });
    return { ...api, errors };
}

const key = load();
const scripting = '{420B2830-E718-11CF-893D-00A0C9054228}';
const build = '16.0.20527.20196';
const base = {
    id: 'sample', description: 'A sample', provenance: 'hand-written', expected: 'observe', mode: 'compile',
    evidencePhase: 'compile', diagnosticMeaning: 'observation', source: 'Sub Main()\nEnd Sub\n',
};
const hash = (kase: object, hostBuild: string | null = build) => key.caseHash(kase, hostBuild);

describe('VBE oracle checkpoint key', () => {
    it.each([
        ['description', 'Another wording'], ['provenance', 'vbe-oracle-verified'], ['expected', 'accepted'],
        ['evidencePhase', 'runtime'], ['diagnosticMeaning', 'compile-valid'],
        ['selectionAfterDialog', { line: 1, column: 1 }], ['crashesHostBuilds', [build]],
    ])('keeps the key when only %s changes, is added or is removed', (field, value) => {
        const withField = { ...base, [field]: value };
        const without: Record<string, unknown> = { ...withField };
        delete without[field];
        expect(hash(withField)).toBe(hash(base));
        expect(hash(without)).toBe(hash(withField));
    });

    it('keeps the key of a case --promote-observed promotes', () => {
        const document = { cases: [structuredClone(base)] };
        const before = hash(document.cases[0]);
        expect(key.promoteObservedCases(document, [{ caseId: 'sample', outcome: 'rejected', stage: 'compile_dialog' }]))
            .toEqual({ promoted: 1, errors: [] });
        expect(document.cases[0]).toMatchObject({ expected: 'rejected', provenance: 'vbe-oracle-verified' });
        expect(hash(document.cases[0])).toBe(before);
    });

    it.each([
        ['source', 'Sub Main()\n    Beep\nEnd Sub\n'], ['mode', 'run'], ['entryPoint', 'Main'], ['host', 'word'],
        ['references', [{ guid: scripting, major: 1, minor: 0 }]],
        ['projectReference', { host: 'excel', modules: [] }],
        ['documentControls', [{ name: 'Button1', progId: 'Forms.CommandButton.1' }]],
        ['modules', [{ name: 'Module1', type: 'standard', source: 'Sub Main()\nEnd Sub\n', entry: true }]],
        ['id', 'renamed'], ['aFieldAddedLater', true],
    ])('changes the key when %s changes', (field, value) => {
        expect(hash({ ...base, [field]: value })).not.toBe(hash(base));
    });

    it('changes the key when anything inside a module changes', () => {
        const module = { name: 'Module1', type: 'standard', source: 'Sub Main()\nEnd Sub\n', entry: true };
        const withModules = { ...base, modules: [module] };
        for (const change of [{ source: "Sub Main()\n'x\nEnd Sub\n" }, { name: 'Module2' }, { type: 'class' },
            { entry: false }, { predeclaredId: true }, { designerControls: [] }]) {
            expect(hash({ ...base, modules: [{ ...module, ...change }] })).not.toBe(hash(withModules));
        }
    });

    it('changes the key with the host build', () => {
        expect(hash(base, '16.0.20527.20197')).not.toBe(hash(base));
        expect(hash(base, null)).not.toBe(hash(base));
    });

    it('leaves out no field the worker or the run path reads', () => {
        const read = new Set([...worker.matchAll(/\$case\.([A-Za-z]+)/gi)].map((match) => match[1]));
        // Every function between the saved result and the worker request.
        for (const name of ['runCaseWithRetries', 'runCaseAttempt', 'runCaseOnce', 'runCaseOnceUnchecked', 'readDialogResult']) {
            for (const match of slice(name).matchAll(/\bkase\.([A-Za-z]+)/g)) { read.add(match[1]); }
        }
        read.delete('PSObject');
        expect([...read].sort()).toEqual(['documentControls', 'entryPoint', 'id', 'mode', 'modules', 'projectReference',
            'references', 'source']);
        expect([...read].filter((field) => key.UNHASHED_FIELDS.has(field))).toEqual([]);
        // runCaseOnceUnchecked hands the host to the worker beside the case.
        expect(key.UNHASHED_FIELDS.has('host')).toBe(false);
    });

    it('leaves out only metadata on every corpus fixture', () => {
        const fields = new Set(cases.flatMap((kase) => Object.keys(kase)));
        expect([...fields].filter((field) => key.UNHASHED_FIELDS.has(field)).sort()).toEqual(['crashesHostBuilds',
            'description', 'diagnosticMeaning', 'evidencePhase', 'expected', 'provenance', 'selectionAfterDialog']);
    });
});

describe('VBE oracle checkpoint file', () => {
    let dir = '';
    afterEach(() => { if (dir) { rmSync(dir, { recursive: true, force: true }); } });
    const write = (lines: object[]) => {
        dir = mkdtempSync(join(tmpdir(), 'xlide-checkpoint-key-'));
        const file = join(dir, 'checkpoint.jsonl');
        writeFileSync(file, lines.map((line) => `${JSON.stringify(line)}\n`).join(''));
        return file;
    };

    it('reuses only results saved under this key format, and names the cases still to run again', () => {
        const api = load();
        const file = write([
            { caseHash: 'old', result: { caseId: 'a', outcome: 'accepted' } },
            { caseHash: 'old', result: { caseId: 'a', outcome: 'accepted' } },
            { key: 1, caseHash: 'old', result: { caseId: 'b', outcome: 'rejected' } },
            { key: '2', caseHash: 'old', result: { caseId: 'f', outcome: 'rejected' } },
            // c ran again under this key: its old line is not counted.
            { caseHash: 'old', result: { caseId: 'c', outcome: 'accepted' } },
            { key: api.CHECKPOINT_KEY_VERSION, caseHash: 'new', result: { caseId: 'c', outcome: 'accepted' } },
            { key: api.CHECKPOINT_KEY_VERSION, caseHash: 'new', result: { caseId: 'd', outcome: 'worker_error' } },
            { caseHash: 'old', result: { caseId: 'e', outcome: 'timeout' } },
        ]);
        const done = api.loadCheckpoint(file);
        expect([...done.keys()]).toEqual([api.checkpointKey('c', 'new')]);
        expect(api.errors).toEqual([`${file}: 4 result(s) for 3 case(s) were saved under another checkpoint key and are not reused.`]);
    });

    it('says nothing when every old case has run again', () => {
        const api = load();
        api.loadCheckpoint(write([{ caseHash: 'old', result: { caseId: 'c', outcome: 'accepted' } },
            { key: api.CHECKPOINT_KEY_VERSION, caseHash: 'h', result: { caseId: 'c', outcome: 'accepted' } }]));
        expect(api.errors).toEqual([]);
    });
});
