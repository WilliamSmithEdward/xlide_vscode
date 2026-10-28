// Issue #1495, measured in Excel, Word, PowerPoint and Access 16.0 (build
// 20527) through the VBE oracle: a recorded reference's name qualifies a
// name in an expression, in any case, bracketed or continued, and stands
// only before a dot. Bare, a library's name is "Expected variable or
// procedure, not project", the host's own library and VBA included. The
// project's own name qualifies its modules, and only it. The cases are the
// issue1495_* fixtures in syntax_corpus/oracle/vbe_oracle_cases.json.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../../src/vbaProjectAnalysis';

type Host = 'excel' | 'word' | 'powerpoint' | 'access';
interface OracleCase {
	id: string;
	expected: 'accepted' | 'rejected';
	host?: Host;
	references?: { guid: string }[];
	modules: { source: string }[];
}

const corpus: { cases: OracleCase[] } = JSON.parse(
	fs.readFileSync(path.join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
);
/** The names a new file of each host records, and its project's name (measured, issue #1413). */
const DEFAULTS: Readonly<Record<Host, readonly string[]>> = {
	excel: ['stdole', 'Office'], word: ['stdole', 'Normal', 'Office'], powerpoint: ['stdole', 'Office'], access: ['stdole', 'DAO'],
};
const PROJECT_NAMES: Readonly<Record<Host, string>> = { excel: 'VBAProject', word: 'Project', powerpoint: 'VBAProject', access: 'Database1' };
const LIBRARY_NAMES: Readonly<Record<string, string>> = {
	'{420B2830-E718-11CF-893D-00A0C9054228}': 'Scripting',
	'{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}': 'DAO',
};

interface Known { referencedLibraries?: string[]; referencedProjects?: string[]; projectName?: string }

function errors(source: string, host: Host = 'excel', known: Known = {}): string[] {
	const project = buildVbaProjectIndex([{ moduleName: 'Module1', source }], undefined, { conditionalCompilation: { projectConstants: {} } });
	return analyzeVbaModuleSource({
		source, moduleName: 'Module1', moduleKind: 'standard', host, ...known,
		...projectAnalysisOptionsForModule(project, 'Module1', projectProcedureSignatures(project)),
	}).diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => `${diagnostic.code}: ${diagnostic.message}`);
}

const measured = corpus.cases.filter((oracleCase) => oracleCase.id.startsWith('issue1495_'));

describe('a reference name as a qualifier (issue #1495)', () => {
	it('has every measured case', () => {
		expect(measured).toHaveLength(61);
	});

	it.each(measured.map((oracleCase) => [oracleCase.id, oracleCase] as const))('%s agrees with the VBE', (_id, oracleCase) => {
		const host = oracleCase.host ?? 'excel';
		const found = errors(oracleCase.modules[0].source, host, {
			referencedLibraries: [...DEFAULTS[host], ...(oracleCase.references ?? []).map((one) => LIBRARY_NAMES[one.guid])],
			projectName: PROJECT_NAMES[host],
		});
		if (oracleCase.expected === 'accepted') {
			expect(found).toEqual([]);
		} else {
			expect(found).toHaveLength(1);
			expect(found[0]).toMatch(/^undeclared-variable: /);
		}
	});

	const module = (...body: string[]): string => ['Option Explicit', 'Public Sub Main()', '    Dim x As Long, v As Variant', ...body.map((line) => `    ${line}`), 'End Sub', ''].join('\r\n');

	it('says a bare library name is a library, as the VBE does', () => {
		expect(errors(module('v = stdole'), 'excel', { referencedLibraries: ['stdole'] })).toEqual([
			"undeclared-variable: 'stdole' names a library or project, which stands only before a dot. This is a VBE compile error: Expected variable or procedure, not project.",
		]);
	});

	it('keeps a variable declared with a library\'s name', () => {
		expect(errors(['Option Explicit', 'Public Sub Main()', '    Dim office As Long', '    office = 1', '    Debug.Print office', 'End Sub', ''].join('\r\n'))).toEqual([]);
	});

	it('leaves the rule as it was where the references were not read', () => {
		expect(errors(module('x = stdole.Checked'))).toHaveLength(1);
		expect(errors(module('x = Excel.xlUp'))).toEqual([]);
	});

	it('does not read a type and the next statement\'s dot as one qualified name', () => {
		// Found by sweeping the accepted corpus with the references known:
		// issue295_08_runtime runs in Excel 16.0.
		const known = { referencedLibraries: ['stdole', 'Office'], projectName: 'VBAProject' };
		expect(errors(['Option Explicit', 'Function Main() As Variant', '    With New Collection', '        .Add 1', '        Main = .Count', '    End With', 'End Function', ''].join('\r\n'), 'excel', known)).toEqual([]);
		expect(errors(['Option Explicit', 'Sub Main()', '    With Application', '        Dim c As Collection', '        .StatusBar = False', '    End With', 'End Sub', ''].join('\r\n'), 'excel', known)).toEqual([]);
		expect(errors(['Option Explicit', 'Sub Main()', '    Dim c As Nowhere _', '        .Thing', 'End Sub', ''].join('\r\n'), 'excel', known).map((e) => e.split(':')[0])).toEqual(['missing-library-reference']);
	});

	it('keeps another VBA project\'s name reported until that project is read', () => {
		// Its module may be Option Private, which makes the call "Method or
		// data member not found"; accepting the name would hide that.
		const known = { referencedLibraries: ['stdole', 'Office', 'ProbeLibrary'], referencedProjects: ['ProbeLibrary'], projectName: 'VBAProject' };
		expect(errors(module('x = ProbeLibrary.Service.Value()'), 'excel', known)).toEqual([
			"undeclared-variable: Variable not defined: 'ProbeLibrary'. Declare it before using it, or remove Option Explicit.",
		]);
	});
});
