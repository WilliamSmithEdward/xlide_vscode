import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { afterAll, describe, expect, it } from 'vitest';
import {
	deleteModule,
	getProtectionInfo,
	getWorkbookInfo,
	listModules,
	listSubs,
	readCells,
	readModule,
	renameModule,
	validateWorkbook,
	writeCells,
	writeModule,
} from '../src/vba/workbookService';
import { openMacroContainer } from '../src/vba/macroContainer';

// Issues #24/#25, engine half: every macro container Office writes, read (and
// written where the write path is sound) natively. Each fixture here was
// authored by the live application itself through pyVBAharness, so the bytes
// are Office ground truth, not synthetic approximations.

const FIXTURES = path.join(__dirname, 'fixtures', 'binaries');
const fixture = (name: string): string => path.join(FIXTURES, name);

const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-containers-'));
afterAll(() => {
	fs.rmSync(tempRoot, { recursive: true, force: true });
});

/** A scratch copy: mutations never touch the checked-in fixture. */
function copyOf(name: string): string {
	const target = path.join(tempRoot, `${Math.random().toString(36).slice(2)}-${name}`);
	fs.copyFileSync(fixture(name), target);
	return target;
}

interface ExpectedModule {
	name: string;
	type: string;
	documentType?: string;
	sourceContains?: string;
}

/** Fixture facts: the modules the authoring script created, per container. */
const READ_CASES: Array<{ file: string; modules: ExpectedModule[] }> = [
	{
		file: 'WordFixture.docm',
		modules: [
			{ name: 'ThisDocument', type: 'document', documentType: 'document', sourceContains: 'Private Sub Document_New()' },
			{ name: 'Module1', type: 'standard', sourceContains: 'Public Sub GreetDocument()' },
			{ name: 'CGreeter', type: 'class', sourceContains: 'Selection.TypeText mGreeting' },
		],
	},
	{
		file: 'WordFixture.doc',
		modules: [
			{ name: 'ThisDocument', type: 'document', documentType: 'document', sourceContains: 'Private Sub Document_New()' },
			{ name: 'Module1', type: 'standard', sourceContains: 'Public Sub GreetDocument()' },
			{ name: 'CGreeter', type: 'class', sourceContains: 'Selection.TypeText mGreeting' },
		],
	},
	{
		file: 'PowerPointFixture.pptm',
		modules: [
			{ name: 'Module1', type: 'standard', sourceContains: 'ActivePresentation.Slides.Count' },
			{ name: 'CDeck', type: 'class', sourceContains: 'ActivePresentation.Name' },
		],
	},
	{
		file: 'PowerPointFixture.ppt',
		modules: [
			// The VBE case-normalized `Count` to `count` in this save; assert
			// the call chain, not the identifier casing it chose that day.
			{ name: 'Module1', type: 'standard', sourceContains: 'ActivePresentation.Slides.' },
			{ name: 'CDeck', type: 'class', sourceContains: 'ActivePresentation.Name' },
		],
	},
	{
		file: 'XlsFixture.xls',
		modules: [
			{ name: 'ThisWorkbook', type: 'document', documentType: 'workbook' },
			{ name: 'Sheet1', type: 'document', documentType: 'worksheet' },
			{ name: 'Module1', type: 'standard', sourceContains: 'Hello from XlsFixture' },
			{ name: 'CGreeter', type: 'class', sourceContains: 'Debug.Print mGreeting' },
		],
	},
	{
		file: 'AccessFixture.accdb',
		modules: [
			{ name: 'Module1', type: 'standard', sourceContains: 'DoCmd.OpenForm "frmGreeting"' },
			{ name: 'CAudit', type: 'class', sourceContains: 'Debug.Print action' },
			{ name: 'MBig', type: 'standard', sourceContains: 'step 800 of the long fixture module' },
		],
	},
	{
		file: 'AccessFixture.mdb',
		modules: [
			{ name: 'Module1', type: 'standard', sourceContains: 'DoCmd.OpenForm "frmGreeting"' },
			{ name: 'CAudit', type: 'class', sourceContains: 'Debug.Print action' },
		],
	},
];

describe.each(READ_CASES)('reading $file', ({ file, modules }) => {
	it('lists the authored modules with their kinds', () => {
		const entries = listModules(fixture(file));
		for (const expected of modules) {
			const entry = entries.find((candidate) => candidate.name === expected.name);
			expect(entry, `${file} should carry ${expected.name}`).toBeDefined();
			expect(entry?.type, `${expected.name} type`).toBe(expected.type);
			if (expected.documentType) {
				expect(entry?.documentType, `${expected.name} documentType`).toBe(expected.documentType);
			}
		}
	});

	it('reads each module source', () => {
		for (const expected of modules) {
			const { source } = readModule(fixture(file), expected.name, true);
			expect(source).toContain(`Attribute VB_Name = "${expected.name}"`);
			if (expected.sourceContains) {
				expect(source).toContain(expected.sourceContains);
			}
		}
	});

	it('answers protection and structural validation', () => {
		const protection = getProtectionInfo(fixture(file));
		expect(protection.isPasswordProtected).toBe(false);
		expect(validateWorkbook(fixture(file)).issues).toEqual([]);
	});
});

describe('the containers behave as their formats require', () => {
	it('finds procedures inside an Access module whose stream chains across LVAL pages', () => {
		const subs = listSubs(fixture('AccessFixture.accdb'), 'MBig');
		expect(subs.map((sub) => sub.name)).toContain('BigTally');
		const { source } = readModule(fixture('AccessFixture.accdb'), 'MBig', true);
		// 800 generated lines force a multi-chunk stream held in a chained
		// long value; every line surviving proves the chain walk.
		for (const step of [1, 400, 800]) {
			expect(source).toContain(`step ${step} of the long fixture module`);
		}
	});

	it('getWorkbookInfo answers modules without a sheet surface for a Word document', () => {
		const info = getWorkbookInfo(fixture('WordFixture.docm'));
		expect(info.sheets).toEqual([]);
		expect(info.namedRanges).toEqual([]);
		expect(info.modules.map((module) => module.name)).toContain('ThisDocument');
	});

	it('refuses sheet reads on containers without worksheets, by name', () => {
		expect(() => readCells(fixture('WordFixture.docm'), 'Sheet1', 'A1'))
			.toThrow(/Word macro-enabled document.*no worksheet surface/);
		expect(() => readCells(fixture('AccessFixture.accdb'), 'Sheet1', 'A1'))
			.toThrow(/Access database.*no worksheet surface/);
		expect(() => readCells(fixture('XlsFixture.xls'), 'Sheet1', 'A1'))
			.toThrow(/legacy Excel workbook.*no worksheet surface/);
	});

	it('classifies containers from content, not extension', () => {
		const disguised = path.join(tempRoot, 'disguised.xlsm');
		fs.copyFileSync(fixture('WordFixture.docm'), disguised);
		const container = openMacroContainer(fs.readFileSync(disguised));
		expect(container.kind).toBe('word');
	});
});

const WRITE_CASES = [
	'WordFixture.docm',
	'PowerPointFixture.pptm',
	'WordFixture.doc',
	'XlsFixture.xls',
	'PowerPointFixture.ppt',
];

describe.each(WRITE_CASES.map((file) => ({ file })))('writing $file', ({ file }) => {
	it('adds a module, preserves the rest, and round-trips through a fresh parse', () => {
		const target = copyOf(file);
		const before = listModules(target).map((module) => module.name).sort();
		const beforeSources = new Map(
			before.map((name) => [name, readModule(target, name, true).source]),
		);

		writeModule(target, 'MAdded', 'Public Sub AddedProc()\r\n    Debug.Print "added"\r\nEnd Sub\r\n');

		const after = listModules(target);
		expect(after.map((module) => module.name)).toContain('MAdded');
		const added = readModule(target, 'MAdded', true).source;
		expect(added).toContain('Attribute VB_Name = "MAdded"');
		expect(added).toContain('Public Sub AddedProc()');
		expect(listSubs(target, 'MAdded').map((sub) => sub.name)).toEqual(['AddedProc']);
		for (const name of before) {
			expect(readModule(target, name, true).source, `${name} must survive the write`)
				.toBe(beforeSources.get(name));
		}
		expect(validateWorkbook(target).issues).toEqual([]);
	});

	it('edits an existing module in place', () => {
		const target = copyOf(file);
		writeModule(target, 'Module1', 'Public Sub Rewritten()\r\nEnd Sub\r\n');
		const { source } = readModule(target, 'Module1', true);
		expect(source).toContain('Public Sub Rewritten()');
		expect(validateWorkbook(target).issues).toEqual([]);
	});
});

describe('rename and delete on a non-Excel host', () => {
	it('round-trips on a Word document', () => {
		const target = copyOf('WordFixture.docm');
		renameModule(target, 'CGreeter', 'CRenamed');
		expect(listModules(target).map((module) => module.name)).toContain('CRenamed');
		expect(readModule(target, 'CRenamed', true).source).toContain('Attribute VB_Name = "CRenamed"');
		deleteModule(target, 'CRenamed');
		expect(listModules(target).map((module) => module.name)).not.toContain('CRenamed');
		expect(validateWorkbook(target).issues).toEqual([]);
	});
});

describe('Access stays read-only, with the reason stated', () => {
	it.each([['AccessFixture.accdb'], ['AccessFixture.mdb']])('%s refuses writes', (file) => {
		const target = copyOf(file);
		expect(() => writeModule(target, 'MNew', 'Public Sub P()\r\nEnd Sub\r\n'))
			.toThrow(/read-only: Access runs VBA from its compiled p-code/);
		expect(fs.readFileSync(target).equals(fs.readFileSync(fixture(file)))).toBe(true);
	});

	it('cell writes name the actual container', () => {
		const target = copyOf('WordFixture.docm');
		expect(() => writeCells(target, 'Sheet1', 'A1', [['x']]))
			.toThrow(/Word macro-enabled document.*cell writes need an OOXML Excel workbook/);
	});
});
