// A name the dir stream records for one of the project's references
// qualifies that library's globals under Option Explicit. Measured in Excel,
// Word, PowerPoint and Access 16.0 (build 20430):
//
// - another VBA project, ProbeLibrary, saved by the host and referenced from
//   a saved consumer project (Tools > References): `ProbeLibrary.Service.
//   Value()`, `ProbeLibrary.Value()`, lower case, brackets, a line
//   continuation before the dot, a call statement and a Debug.Print argument
//   compile; `Main = ProbeLibrary` is "Expected variable or procedure, not
//   project", and a qualifier no reference names is "Variable not defined".
//   The oracle corpus cannot hold these (it adds registered libraries only),
//   so the measured sources are here;
// - registered libraries, in syntax_corpus/oracle/vbe_oracle_cases.json
//   (library_qualifier_*): `Scripting.BinaryCompare` compiles with the
//   reference and is "Variable not defined" without it, and `Main =
//   Scripting` is "Expected variable or procedure, not project". Libraries
//   no host model carries go by their recorded name alone: `stdole.Checked`
//   and, with the DAO reference, `DAO.dbOpenDynaset` compile, `Main =
//   stdole` does not, and without the reference `DAO.dbOpenDynaset` is
//   "Variable not defined". The corpus sweep analyzes a case without the
//   recorded names, so these are measured here only.

import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('../../helpers/vscodeMock')).vscodeMock());

import { AnalysisWorkerState } from '../../../src/analysisWorkerLogic';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import { analyzeProject, resetProjectAnalysisResultCacheForTests } from '../../../src/vbaProjectWideAnalysis';
import { fakeProjectEngine } from '../../helpers/fakeProjectEngine';

type Host = 'excel' | 'word' | 'powerpoint' | 'access';
const HOSTS: readonly Host[] = ['excel', 'word', 'powerpoint', 'access'];
/** What each measured consumer's dir stream records (read from the saved files). */
const RECORDED: Readonly<Record<Host, readonly string[]>> = {
	excel: ['stdole', 'Office', 'ProbeLibrary'],
	word: ['stdole', 'Office', 'ProbeLibrary'],
	powerpoint: ['stdole', 'Office', 'ProbeLibrary'],
	access: ['stdole', 'DAO', 'ProbeLibrary'],
};
const main = (body: string): string => `Option Explicit\nPublic Function Main() As Variant\n${body}End Function\n`;

/** Each error as `code: spanned text`. */
function errors(source: string, host: Host, referencedLibraries: readonly string[] | undefined): string[] {
	return analyzeVbaModuleSource({
		source,
		moduleName: 'Module1',
		moduleKind: 'standard',
		host,
		knownIdentifiers: new Set(['main']),
		...(referencedLibraries ? { referencedLibraries } : {}),
	}).diagnostics
		.filter((diagnostic) => diagnostic.severity === 'error')
		.map((diagnostic) => `${diagnostic.code}: ${source.slice(diagnostic.span.start, diagnostic.span.end)}`);
}

const COMPILES: Readonly<Record<string, string>> = {
	qualified: 'Main = ProbeLibrary.Service.Value()\n',
	project_procedure: 'Main = ProbeLibrary.Value()\n',
	lowercase: 'Main = probelibrary.Service.Value()\n',
	bracketed: 'Main = [ProbeLibrary].Service.Value()\n',
	continued: 'Main = ProbeLibrary _\n.Service.Value()\n',
	call_statement: 'ProbeLibrary.Service.Value\nMain = 1\n',
	print: 'Debug.Print ProbeLibrary.Service.Value()\nMain = 1\n',
};

describe('a referenced project qualifies its globals', () => {
	for (const host of HOSTS) {
		for (const [shape, body] of Object.entries(COMPILES)) {
			it(`${shape} compiles in ${host}`, () => {
				expect(errors(main(body), host, RECORDED[host])).toEqual([]);
			});
		}
		it(`the project name alone is no value in ${host}`, () => {
			expect(errors(main('Main = ProbeLibrary\n'), host, RECORDED[host])).toEqual(['undeclared-variable: ProbeLibrary']);
		});
		it(`a qualifier no reference names is still undeclared in ${host}`, () => {
			expect(errors(main('Main = OtherLibrary.Service.Value()\n'), host, RECORDED[host])).toEqual(['undeclared-variable: OtherLibrary']);
		});
		it(`the qualifier is undeclared where the references were not read in ${host}`, () => {
			expect(errors(main('Main = ProbeLibrary.Service.Value()\n'), host, undefined)).toEqual(['undeclared-variable: ProbeLibrary']);
		});
	}

	it('matches the recorded name in any case', () => {
		expect(errors(main('Main = PROBELIBRARY.Service.Value()\n'), 'excel', ['stdole', 'Office', 'probeLibrary'])).toEqual([]);
	});

	it('is not member access through `!`', () => {
		expect(errors(main('Main = ProbeLibrary!Service\n'), 'excel', RECORDED.excel)).toEqual(['undeclared-variable: ProbeLibrary']);
	});

	it('still checks the arguments of a qualified call', () => {
		expect(errors(main('Main = ProbeLibrary.Service.Value(missing)\n'), 'excel', RECORDED.excel)).toEqual(['undeclared-variable: missing']);
	});

	it('an empty reference list qualifies nothing', () => {
		expect(errors(main('Main = ProbeLibrary.Service.Value()\n'), 'excel', [])).toEqual(['undeclared-variable: ProbeLibrary']);
	});
});

/** What a new file of each host records before any reference is added (measured). */
const DEFAULTS: Readonly<Record<Host, readonly string[]>> = {
	excel: ['stdole', 'Office'], word: ['stdole', 'Normal', 'Office'], powerpoint: ['stdole', 'Office'], access: ['stdole', 'DAO'],
};

describe('a library no host model carries qualifies by its recorded name', () => {
	for (const host of HOSTS) {
		it(`stdole.Checked compiles in ${host}`, () => {
			expect(errors(main('Main = stdole.Checked\n'), host, DEFAULTS[host])).toEqual([]);
		});
		it(`stdole alone is no value in ${host}`, () => {
			expect(errors(main('Main = stdole\n'), host, DEFAULTS[host])).toEqual(['undeclared-variable: stdole']);
		});
	}
	for (const host of ['excel', 'word', 'powerpoint'] as const) {
		it(`DAO.dbOpenDynaset compiles with the reference in ${host}`, () => {
			expect(errors(main('Main = DAO.dbOpenDynaset\n'), host, [...DEFAULTS[host], 'DAO'])).toEqual([]);
		});
		it(`DAO.dbOpenDynaset is undeclared without it in ${host}`, () => {
			expect(errors(main('Main = DAO.dbOpenDynaset\n'), host, DEFAULTS[host])).toEqual(['undeclared-variable: DAO']);
		});
	}
});

describe('the reference list reaches the rule', () => {
	it('from a project the engine reads', async () => {
		resetProjectAnalysisResultCacheForTests();
		const engine = fakeProjectEngine([{
			name: 'Module1',
			type: 'standard',
			source: main('Main = ProbeLibrary.Service.Value()\n'),
			projectReferences: [
				{ name: 'stdole', kind: 'registered', libid: '*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation' },
				{ name: 'ProbeLibrary', kind: 'project', libid: '*\\CC:\\Work\\Provider.xlsm' },
			],
		}]);
		const result = await analyzeProject(engine, 'Consumer.xlsm');
		expect(result.problems.filter((problem) => problem.code === 'undeclared-variable')).toEqual([]);
		resetProjectAnalysisResultCacheForTests();
	});

	it('through the worker, when only the references change', () => {
		const state = new AnalysisWorkerState();
		const source = main('Main = ProbeLibrary.Service.Value()\n');
		state.handle({ kind: 'seed', projectKey: 'refs', generation: 1,
			modules: [{ moduleName: 'Module1', source, type: 'standard' }] });
		let requestId = 0;
		const codes = (referencedLibraries: readonly string[]): string[] => {
			const response = state.handle({ kind: 'analyze', requestId: ++requestId, docKey: 'doc', projectKey: 'refs',
				generation: 1, source, moduleName: 'Module1', moduleType: 'standard', host: 'excel', referencedLibraries });
			if (response?.kind !== 'result') { throw new Error('Expected worker result'); }
			return response.diagnostics.map((diagnostic) => diagnostic.code ?? '');
		};
		expect(codes(['stdole', 'Office'])).toContain('undeclared-variable');
		expect(codes(['stdole', 'Office', 'ProbeLibrary'])).not.toContain('undeclared-variable');
		expect(codes(['stdole', 'Office'])).toContain('undeclared-variable');
	});
});
