// member-exists-in-base-class: a module-level declaration in a module that
// derives from a host class takes a name that class already has. Measured in
// Excel, Word, PowerPoint and Access 16.0 (2026-10-16) through the VBE
// oracle; the base_member_* fixtures in vbe_oracle_cases.json hold every
// verdict, and this test holds the rule to each one.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../../../src/analyzer';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import { byCode } from '../../helpers/diagnostics';

const CODE = 'member-exists-in-base-class';

interface Fixture {
	id: string;
	host: string;
	expected: 'accepted' | 'rejected';
	description: string;
	modules: Array<{ name: string; type: string; source: string }>;
}

const fixtures: Fixture[] = (JSON.parse(readFileSync(join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8')) as { cases: Fixture[] })
	.cases.filter((kase) => kase.id.startsWith('base_member_'));

/** What the extension tells the analyzer about each kind of module. */
function moduleOptions(host: string, mod: Fixture['modules'][number]): Record<string, unknown> {
	if (host === 'access' && mod.type === 'document') {
		const report = mod.name.startsWith('Report_');
		return { moduleKind: 'userform', designerClass: report ? 'Access.Report' : 'Access.Form' };
	}
	if (mod.type === 'document' && host === 'excel' && mod.name !== 'ThisWorkbook') {
		return { moduleKind: 'document', documentType: 'worksheet' };
	}
	return { moduleKind: mod.type === 'userform' ? 'userform' : mod.type === 'document' ? 'document' : 'standard' };
}

function clashes(host: string, mod: Fixture['modules'][number]): number {
	const { diagnostics } = analyzeVbaModuleSource({
		source: mod.source,
		moduleName: mod.name,
		host,
		...moduleOptions(host, mod),
	} as Parameters<typeof analyzeVbaModuleSource>[0]);
	return diagnostics.filter((diagnostic) => diagnostic.code === CODE).length;
}

describe('member-exists-in-base-class against the measured fixtures', () => {
	it('has fixtures for every Office host', () => {
		expect(new Set(fixtures.map((fixture) => fixture.host))).toEqual(new Set(['excel', 'word', 'powerpoint', 'access']));
	});

	it.each(fixtures.map((fixture) => [fixture.id, fixture] as const))('%s', (_id, fixture) => {
		const clash = fixture.expected === 'rejected' && fixture.description.includes('Member already exists');
		const found = fixture.modules.reduce((total, mod) => total + clashes(fixture.host, mod), 0);
		expect(found).toBe(clash ? 1 : 0);
	});
});

describe('member-exists-in-base-class elsewhere', () => {
	const words = 'Public Function Words() As Long\n    Words = 1\nEnd Function\n';

	it('names the class and the clash', () => {
		const [diagnostic] = byCode(analyzeModule(words, { host: 'word', moduleName: 'ThisDocument', moduleKind: 'document' }), CODE);
		expect(diagnostic?.message).toContain("'Words' is already a member of Word.Document");
		expect(diagnostic?.message).toContain('Member already exists in an object module from which this object module derives');
	});

	it('holds a sheet of unknown type only to what a Worksheet and a Chart share', () => {
		const sheet = (src: string) => byCode(analyzeModule(src, { moduleName: 'Sheet1', moduleKind: 'document' }), CODE);
		// Name is on both; Range is a Worksheet's only, and the sheet may be a chart.
		expect(sheet('Public Function Name() As String\nEnd Function\n')).toHaveLength(1);
		expect(sheet('Public Sub Range()\nEnd Sub\n')).toHaveLength(0);
	});

	it('leaves standard modules, class modules and VB6 designers alone', () => {
		expect(byCode(analyzeModule(words, { host: 'word', moduleName: 'Module1', moduleKind: 'standard' }), CODE)).toHaveLength(0);
		expect(byCode(analyzeModule(words, { host: 'word', moduleName: 'Class1', moduleKind: 'class' }), CODE)).toHaveLength(0);
		const caption = 'Public Function Caption() As String\nEnd Function\n';
		expect(byCode(analyzeModule(caption, { host: 'vb6', moduleName: 'Form1', moduleKind: 'userform', designerClass: 'VB.Form' }), CODE)).toHaveLength(0);
	});

	it('does not report a Public Const, Declare or array, which an object module refuses for its own reason', () => {
		const src = [
			'Public Const Words As Long = 1',
			'Public Declare PtrSafe Function Words Lib "kernel32" Alias "GetTickCount" () As Long',
			'Public Words() As Long',
			'',
		].join('\n');
		expect(byCode(analyzeModule(src, { host: 'word', moduleName: 'ThisDocument', moduleKind: 'document' }), CODE)).toHaveLength(0);
	});
});
