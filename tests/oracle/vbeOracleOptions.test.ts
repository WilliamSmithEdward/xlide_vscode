// The VBE oracle coordinator's option values, fixture hosts and checkpoint
// paths (issue #1440), checked before anything runs.

import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { runInNewContext } from 'node:vm';
import { afterEach, describe, expect, it } from 'vitest';
import { OFFICE_HOST_APPS } from '../../src/officeHostApps';

const source = fs.readFileSync(path.join(process.cwd(), 'syntax_corpus/oracle/run_vbe_oracle.mjs'), 'utf8');
const lift = (name: string): string => {
	const start = source.indexOf(`function ${name}(`);
	return source.slice(start, source.indexOf('\n}', start) + 2);
};
const api = runInNewContext(`
	const DEFAULT_HOST = 'excel';
	const DEFAULT_CASES = 'vbe_oracle_cases.json';
	const DEFAULT_TIMEOUT_RETRIES = 2;
	const DEFAULT_CASES_PER_INSTANCE = 40;
	const OFFICE_HOSTS = ${JSON.stringify(Object.fromEntries(Object.keys(OFFICE_HOST_APPS).map((host) => [host, {}])))};
	${['parseArgs', 'caseHost', 'unsupportedHost', 'malformedReferences', 'unwritableCheckpoint'].map(lift).join('\n')}
	({ parseArgs, caseHost, unsupportedHost, malformedReferences, unwritableCheckpoint })`, { fs, path }) as {
	parseArgs: (argv: string[]) => Record<string, unknown> & { caseIds: string[] };
	caseHost: (kase: object) => string;
	unsupportedHost: (kase: object) => string | undefined;
	malformedReferences: (kase: object) => string | undefined;
	unwritableCheckpoint: (file: string) => string | undefined;
};

const VALUE_OPTIONS = ['--cases', '--case', '--checkpoint', '--timeout', '--timeout-retries', '--cases-per-instance', '--cases-per-excel', '--dialog-hold-seconds'];

describe('oracle option values', () => {
	it.each(VALUE_OPTIONS)('refuses %s with no value', (flag) => {
		expect(() => api.parseArgs([flag])).toThrow(`${flag} expects a value`);
		expect(() => api.parseArgs([flag, ''])).toThrow(`${flag} expects a value`);
	});

	it.each(VALUE_OPTIONS)('refuses another option in place of the value of %s', (flag) => {
		for (const next of ['--strict', '--json', '--case', '--help']) {
			expect(() => api.parseArgs([flag, next, 'x'])).toThrow(`${flag} expects a value, not the option '${next}'`);
		}
	});

	it('refuses an integer option given anything but an integer', () => {
		for (const raw of ['5abc', '1.5', 'oops', '0x10', '1e3', ' 5']) {
			expect(() => api.parseArgs(['--timeout', raw]), raw).toThrow(`--timeout expects an integer (got '${raw}')`);
		}
		expect(api.parseArgs(['--timeout-retries', '-1']).timeoutRetries).toBe(-1);
	});

	it('keeps repeated cases, fixture paths and checkpoint paths, spaces included', () => {
		const args = api.parseArgs(['--case', 'first', '--case', 'second case', '--cases', 'C:\\My Fixtures\\cases.json',
			'--checkpoint', 'D:\\run results\\checkpoint.jsonl', '--strict', '--json', '--timeout', '30']);
		expect(args.caseIds).toEqual(['first', 'second case']);
		expect(args).toMatchObject({ cases: 'C:\\My Fixtures\\cases.json', checkpoint: 'D:\\run results\\checkpoint.jsonl', strict: true, json: true, timeout: 30 });
	});

	it('takes a value that begins with one dash, and a path written ./--name', () => {
		expect(api.parseArgs(['--case', '-dash-id']).caseIds).toEqual(['-dash-id']);
		expect(api.parseArgs(['--checkpoint', './--odd.jsonl']).checkpoint).toBe('./--odd.jsonl');
	});
});

describe('oracle fixture hosts', () => {
	it.each(Object.keys(OFFICE_HOST_APPS))('takes %s, and Excel for a case with no host', (host) => {
		expect(api.unsupportedHost({ id: 'c', host })).toBeUndefined();
		expect(api.caseHost({ id: 'c', host })).toBe(host);
		expect(api.unsupportedHost({ id: 'c' })).toBeUndefined();
		expect(api.caseHost({ id: 'c' })).toBe('excel');
	});

	it('refuses every other value the same way, naming the case', () => {
		for (const host of ['Excel', 'WORD', ' excel', 'excel ', '', 'outlook', 'constructor', null, 1, true, {}, ['excel']]) {
			expect(api.unsupportedHost({ id: 'the-case', host }), JSON.stringify(host)).toBe(`the-case names host ${JSON.stringify(host)}`);
		}
	});
});

describe('oracle fixture references', () => {
	const OUTLOOK = { guid: '{00062FFF-0000-0000-C000-000000000046}', major: 9, minor: 6 };

	it('takes none, an empty list, and libraries named by GUID and version', () => {
		expect(api.malformedReferences({ id: 'c' })).toBeUndefined();
		expect(api.malformedReferences({ id: 'c', references: [] })).toBeUndefined();
		expect(api.malformedReferences({ id: 'c', references: [OUTLOOK, { guid: '{420b2830-e718-11cf-893d-00a0c9054228}', major: 1, minor: 0 }] })).toBeUndefined();
	});

	it('refuses anything else, naming the case', () => {
		for (const references of [OUTLOOK, 'Outlook', null, [null], [{ ...OUTLOOK, guid: '00062FFF-0000-0000-C000-000000000046' }],
			[{ ...OUTLOOK, guid: '{00062FFF}' }], [{ ...OUTLOOK, major: '9' }], [{ ...OUTLOOK, minor: 1.5 }], [{ ...OUTLOOK, major: -1 }], [{ guid: OUTLOOK.guid }]]) {
			expect(api.malformedReferences({ id: 'the-case', references }), JSON.stringify(references))
				.toBe(`the-case has references ${JSON.stringify(references)}`);
		}
	});

	it('is checked before Office starts, and the worker adds each by GUID before the modules', () => {
		const worker = fs.readFileSync(path.join(process.cwd(), 'syntax_corpus/oracle/vbe_oracle_worker.ps1'), 'utf8');
		const add = worker.indexOf('$vbProject.References.AddFromGuid([string]$reference.guid, [int]$reference.major, [int]$reference.minor)');
		expect(add).toBeGreaterThan(0);
		expect(add).toBeLessThan(worker.indexOf('Set-Stage "add_modules"'));
		expect(source.indexOf('cases.map(malformedReferences)')).toBeLessThan(source.indexOf('const hosts = [...new Set(cases.map(caseHost))];'));
	});
});

describe('oracle checkpoint paths', () => {
	const made: string[] = [];
	afterEach(() => {
		for (const dir of made.splice(0)) {
			fs.rmSync(dir, { recursive: true, force: true });
		}
	});
	const tempDir = (): string => {
		const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-oracle-checkpoint-'));
		made.push(dir);
		return dir;
	};

	it('takes a new file in an existing directory, and an existing file, writing neither', () => {
		const dir = tempDir();
		expect(api.unwritableCheckpoint(path.join(dir, 'new checkpoint.jsonl'))).toBeUndefined();
		expect(fs.existsSync(path.join(dir, 'new checkpoint.jsonl'))).toBe(false);
		const existing = path.join(dir, 'existing.jsonl');
		fs.writeFileSync(existing, '{"x":1}\n');
		expect(api.unwritableCheckpoint(existing)).toBeUndefined();
		expect(fs.readFileSync(existing, 'utf8')).toBe('{"x":1}\n');
	});

	it('refuses a directory, and a file whose directory does not exist', () => {
		const dir = tempDir();
		expect(api.unwritableCheckpoint(dir)).toBe('is a directory');
		expect(api.unwritableCheckpoint(path.join(dir, 'missing', 'checkpoint.jsonl'))).toContain('does not exist');
	});
});
