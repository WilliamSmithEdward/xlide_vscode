// Issue #1419, measured in Excel, Word, PowerPoint and Access 16.0 (build
// 20430) through the VBE oracle, compile-only: UserForm designer controls
// and a predeclared class that take a VBA function's name, a worksheet's or
// document's unknown member in block headers, and keyword-named VBA
// functions written bare. The fixtures carry what the designer and the
// attributes declare, and the analyzer reads them through the same adapter
// the corpus sweeps use. The cases are in syntax_corpus/oracle/vbe_oracle_cases.json.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { oracleModuleInput, type OracleFixtureModule } from '../../helpers/oracleModuleInput';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import {
	buildVbaProjectIndex,
	effectiveModuleKind,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
} from '../../../src/vbaProjectAnalysis';

interface OracleCase {
	id: string;
	description: string;
	expected: 'accepted' | 'rejected';
	host?: 'excel' | 'word' | 'powerpoint' | 'access';
	modules: OracleFixtureModule[];
	selectionAfterDialog?: { module: string; startLine: number; startColumn: number; endLine: number; endColumn: number };
}

const corpus: { cases: OracleCase[] } = JSON.parse(
	readFileSync(join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
);
const cases = corpus.cases.filter((oracleCase) => oracleCase.id.startsWith('issue1419_'));
const family = (id: string): string => id.replace(/^issue1419_(word_|powerpoint_|access_)?/, '').replace(/_compile$/, '');

/** The analyzer's code for each compile error the VBE gave. */
const CODE_FOR_MESSAGE: Record<string, string> = {
	'Type mismatch': 'is-operator-non-object',
	'Syntax error': 'reserved-keyword-in-expression',
	'Method or data member not found': 'member-not-found',
};

/**
 * `Len() Is Nothing` is the VBE's Syntax error, reported as argument-count.
 * `Me.Len` on a form whose fixture does not state its control list stays
 * quiet: without the list nothing proves Len absent. The same form with an
 * empty designer, u_empty_designer_me_len, is reported.
 */
const OTHER_CODES: Record<string, string[]> = {
	u_len_control_4: ['argument-count'],
	u_no_control_me_len: [],
};

function errorCodes(oracleCase: OracleCase): string[] {
	const modules = oracleCase.modules.map(oracleModuleInput);
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	return modules.flatMap((mod) => analyzeVbaModuleSource({
		source: mod.source,
		moduleName: mod.moduleName,
		moduleKind: effectiveModuleKind(mod),
		host: oracleCase.host ?? 'excel',
		...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
	}).diagnostics
		.filter((diagnostic) => diagnostic.severity === 'error')
		.map((diagnostic) => diagnostic.code));
}

describe('designer, attribute and document inputs (#1419)', () => {
	it('holds every measured case', () => {
		expect(cases).toHaveLength(233);
		expect(new Set(cases.map((oracleCase) => oracleCase.host ?? 'excel'))).toEqual(new Set(['excel', 'word', 'powerpoint', 'access']));
	});

	it.each(cases.map((oracleCase): [string, OracleCase] => [oracleCase.id, oracleCase]))('%s gets the error the VBE gave', (_id, oracleCase) => {
		const message = / is rejected: Compile error: (.+?)\. Measured in /.exec(oracleCase.description)?.[1];
		const expected = OTHER_CODES[family(oracleCase.id)]
			?? (oracleCase.expected === 'accepted' ? [] : [CODE_FOR_MESSAGE[message ?? '']]);
		expect(expected).not.toContain(undefined);
		expect(errorCodes(oracleCase)).toEqual(expected);
	});

	it('keeps the VBE selection after each compile dialog, on every rejection', () => {
		for (const oracleCase of cases) {
			const selection = oracleCase.selectionAfterDialog;
			if (oracleCase.expected === 'accepted') {
				expect(selection, oracleCase.id).toBeUndefined();
				continue;
			}
			expect(selection, oracleCase.id).toBeDefined();
			expect(oracleCase.modules.map((mod) => mod.name)).toContain(selection!.module);
			expect(selection!.startLine).toBeGreaterThan(0);
			expect(selection!.endLine).toBeGreaterThanOrEqual(selection!.startLine);
		}
	});
});

describe('the fixture adapter', () => {
	it('types each designer control by its ProgID', () => {
		expect(oracleModuleInput({ name: 'Form1', type: 'userform', source: '', designerControls: [{ name: 'Picker', progId: 'Forms.ComboBox.1' }] }))
			.toEqual({ moduleName: 'Form1', type: 'userform', source: '', implicitMembers: [{ name: 'Picker', type: 'MSForms.ComboBox' }] });
	});

	it('makes an empty designer an empty, known control list', () => {
		expect(oracleModuleInput({ name: 'Form1', type: 'userform', source: '', designerControls: [] }).implicitMembers).toEqual([]);
		expect(oracleModuleInput({ name: 'Form1', type: 'userform', source: '' }).implicitMembers).toBeUndefined();
	});

	it('carries a document module\'s designer class and a class\'s PredeclaredId', () => {
		expect(oracleModuleInput({ name: 'Sheet1', type: 'document', source: '', designerClass: 'Excel.Worksheet', implicitMembers: [] }))
			.toEqual({ moduleName: 'Sheet1', type: 'document', source: '', implicitMembers: [], designerClass: 'Excel.Worksheet' });
		expect(oracleModuleInput({ name: 'Timer', type: 'class', source: '', predeclaredId: true }).predeclaredId).toBe(true);
	});
});
