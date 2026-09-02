import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import {
	compileExpectation,
	modulesOf,
	splitOptions,
	stageCase,
	twinModule,
} from '../syntax_corpus/oracle/twinbasic/run_twinbasic_oracle.mjs';

// The twinBASIC oracle's staging (roadmap_vb6_support.md, Slice 4): what a
// VBE corpus case becomes on disk before the compiler sees it. Pinned here
// because the verdict depends on it: a procedure the startup module does
// not reach is never compiled (measured), so the touch sub and the entry
// call decide whether a rejection can be seen at all. Nothing here runs
// twinBASIC.

const roots: string[] = [];
afterEach(() => {
	for (const root of roots.splice(0)) { fs.rmSync(root, { recursive: true, force: true }); }
});

function stageRoot(): string {
	const root = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-oracle-staging-test-'));
	roots.push(root);
	return root;
}

describe('staging a corpus case as a twinBASIC project', () => {
	it('wraps a standard module and reaches every procedure by AddressOf', () => {
		const { text, procedures, touch } = twinModule('XlideOracleModule', 'standard',
			'Option Explicit\n\nPublic Sub XlideOracleEntry()\nEnd Sub\n\nPrivate Function Helper(ByVal n As Long) As Long\n    Helper = n\nEnd Function\n');
		expect(text.startsWith('Module XlideOracleModule\r\n    Option Explicit\r\n')).toBe(true);
		expect(text).toContain('Public Sub XlideOracleEntry()');
		expect(procedures).toEqual(['XlideOracleEntry', 'Helper']);
		expect(touch).toBe('XlideOracleTouch_XlideOracleModule');
		expect(text).toContain('        p = AddressOf XlideOracleEntry\r\n        p = AddressOf Helper\r\n');
		expect(text.trimEnd().endsWith('End Module')).toBe(true);
	});

	it('wraps a class module without a touch sub', () => {
		const { text, touch } = twinModule('Person', 'class', 'Public Name As String\n\nPublic Function Greet() As String\n    Greet = "hi " & Name\nEnd Function\n');
		expect(text.startsWith('Class Person\r\n')).toBe(true);
		expect(text).not.toContain('AddressOf');
		expect(touch).toBeUndefined();
	});

	it('hoists every Option line and leaves the body otherwise alone', () => {
		const { options, body } = splitOptions('Option Explicit\nOption Base 1\nSub T()\nEnd Sub\n');
		expect(options).toEqual(['Option Explicit', 'Option Base 1']);
		expect(body.trim()).toBe('Sub T()\nEnd Sub');
	});

	it('reads both corpus shapes', () => {
		expect(modulesOf({ id: 'a', source: 'Sub T()\nEnd Sub\n' })).toEqual([
			{ name: 'XlideOracleModule', type: 'standard', source: 'Sub T()\nEnd Sub\n', entry: true },
		]);
		expect(modulesOf({ id: 'b', modules: [{ name: 'Shared', type: 'standard', source: 'x' }, { name: 'Entry', type: 'standard', entry: true, source: 'y' }] }))
			.toEqual([
				{ name: 'Shared', type: 'standard', source: 'x', entry: false },
				{ name: 'Entry', type: 'standard', source: 'y', entry: true },
			]);
	});

	it('writes the project folder with a startup that calls the touch subs and the entry', () => {
		const root = stageRoot();
		const stage = stageCase({
			id: 'runtime_probe',
			entryPoint: 'XlideOracleEntry',
			modules: [
				{ name: 'SharedRuntimeArgs', type: 'standard', source: 'Public Const SharedBadLength As Long = -1\n' },
				{ name: 'XlideOracleModule', type: 'standard', entry: true, source: 'Public Sub XlideOracleEntry()\nEnd Sub\n' },
				{ name: 'Person', type: 'class', source: 'Public Name As String\n' },
			],
		}, root);
		expect(fs.existsSync(path.join(stage.dir, 'Settings'))).toBe(true);
		expect(JSON.parse(fs.readFileSync(path.join(stage.dir, 'Settings'), 'utf8'))['project.optionExplicit']).toBe(false);
		const sources = fs.readdirSync(path.join(stage.dir, 'Sources')).sort();
		expect(sources).toEqual(['Person.twin', 'SharedRuntimeArgs.twin', 'XlideOracleModule.twin', 'XlideOracleStartup.twin']);
		const startup = fs.readFileSync(path.join(stage.dir, 'Sources', 'XlideOracleStartup.twin'), 'utf8');
		expect(startup).toContain('SharedRuntimeArgs.XlideOracleTouch_SharedRuntimeArgs');
		expect(startup).toContain('XlideOracleModule.XlideOracleTouch_XlideOracleModule');
		expect(startup).toContain('Dim PersonInstance As New Person');
		expect(startup).toContain('XlideOracleModule.XlideOracleEntry');
		expect(stage.reach).toEqual([
			'SharedRuntimeArgs (0 procedures by AddressOf)',
			'XlideOracleModule (1 procedure by AddressOf)',
			'Person (instantiated)',
		]);
		expect(stage.definesMain).toBe(false);
	});

	it('flags a case that declares its own Sub Main', () => {
		const stage = stageCase({ id: 'has_main', source: 'Sub Main()\nEnd Sub\n' }, stageRoot());
		expect(stage.definesMain).toBe(true);
	});

	it('compares a runtime case as accepted at compile time', () => {
		expect(compileExpectation({ expected: 'rejected', mode: 'run', evidencePhase: 'runtime' })).toBe('accepted');
		expect(compileExpectation({ expected: 'rejected', mode: 'compile', evidencePhase: 'compile' })).toBe('rejected');
		expect(compileExpectation({ expected: 'accepted', mode: 'compile' })).toBe('accepted');
		expect(compileExpectation({})).toBe('observe');
	});
});
