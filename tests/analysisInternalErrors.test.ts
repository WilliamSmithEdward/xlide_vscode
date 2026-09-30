import { afterEach, describe, expect, it } from 'vitest';
import { analyzeModule, type AnalysisFailure, type AnalyzeModuleOptions } from '../src/analyzer/diagnostics/analyzeModule';
import { DIAGNOSTIC_RULE_REGISTRY, type DiagnosticRuleEntry } from '../src/analyzer/diagnostics/registry';
import { ProjectIndex } from '../src/analyzer/symbols/projectIndex';
import { buildLiveVbaProjectIndex, projectAnalysisOptionsForModule } from '../src/vbaProjectAnalysis';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';
import { AnalysisWorkerState } from '../src/analysisWorkerLogic';
import { logAnalysisFailures, setAnalysisFailureLog } from '../src/analysisFailureLog';

// analyzeModule never throws, and until issue #178 a failure of its own was
// invisible: a rule that threw, or one walk visitor that threw and stopped
// the walk for every rule after it, came back as fewer findings with nothing
// to say so. Each failure is now reported to onInternalError, and a visitor
// that throws stops only its own rule.

const LIB = 'Option Explicit\r\nPublic Function Twice(ByVal n As Long) As Long\r\n    Twice = n * 2\r\nEnd Function\r\n';
const CALLER = [
	'Option Explicit',
	'Sub First()',
	'    Dim r As Long',
	'    r = Twice(1, 2)',
	'End Sub',
	'Sub Second()',
	'    Dim q As Long',
	'    q = Twice(3, 4)',
	'End Sub',
	'',
].join('\r\n');

function project(): { index: ProjectIndex; options: AnalyzeModuleOptions } {
	const index = new ProjectIndex();
	index.setModule({ moduleName: 'Lib', moduleKind: 'standard', source: LIB });
	index.setModule({ moduleName: 'Caller', moduleKind: 'standard', source: CALLER });
	const options = { ...projectAnalysisOptionsForModule(index, 'Caller'), moduleName: 'Caller', moduleKind: 'standard' as const };
	return { index, options };
}

function analyze(options: AnalyzeModuleOptions): { codes: string[]; failures: Array<AnalysisFailure & { message: string }> } {
	const failures: Array<AnalysisFailure & { message: string }> = [];
	const codes = analyzeModule(CALLER, {
		...options,
		onInternalError: (error, where) => failures.push({ ...where, message: (error as Error).message }),
	}).map((diagnostic) => `${diagnostic.code}@${diagnostic.span.start}`);
	return { codes, failures };
}

const registry = DIAGNOSTIC_RULE_REGISTRY as DiagnosticRuleEntry[];

/** Runs with a rule added to the registry for the length of one analysis. */
function withRule<T>(rule: DiagnosticRuleEntry, run: () => T): T {
	registry.push(rule);
	try {
		return run();
	} finally {
		registry.splice(registry.indexOf(rule), 1);
	}
}

describe('analyzeModule reports what it could not check (issue #178)', () => {
	afterEach(() => {
		expect(registry.some((rule) => rule.name.startsWith('probe'))).toBe(false);
	});

	it('reports the array form of projectProcedures, where nine rules used to throw unseen', () => {
		const { index, options } = project();
		const withMap = analyze(options);
		expect(withMap.failures).toEqual([]);
		expect(withMap.codes.filter((code) => code.startsWith('argument-count'))).toHaveLength(2);

		const withArray = analyze({ ...options, projectProcedures: index.visibleProcedureSignatures('Caller') as never });
		expect(withArray.failures).toEqual([{
			stage: 'options',
			message: 'projectProcedures must be a Map of lowercased name to signatures (projectProcedureSignatures), not an array',
		}]);
		// Left out, not thrown on: every rule runs, as with no project procedures.
		expect(withArray.codes).toEqual(analyze({ ...options, projectProcedures: undefined }).codes);
	});

	it('reports a rule that throws, and keeps every other rule s findings', () => {
		const { options } = project();
		const baseline = analyze(options);
		const probed = withRule({ name: 'probeRun', run: () => { throw new Error('run failed'); } }, () => analyze(options));
		expect(probed.failures).toEqual([{ stage: 'rule', rule: 'probeRun', message: 'run failed' }]);
		expect(probed.codes).toEqual(baseline.codes);
	});

	it('stops only the statement visitor that throws, once, and the walk goes on for the rest', () => {
		const { options } = project();
		const baseline = analyze(options);
		expect(baseline.codes.length).toBeGreaterThan(0);
		// Registered first, it walks ahead of every real rule: its throw used to
		// end the walk there, for all of them and every procedure after.
		const probe: DiagnosticRuleEntry = {
			name: 'probeStatements',
			procedureStatements: () => () => { throw new Error('statement visitor failed'); },
		};
		registry.unshift(probe);
		let probed: ReturnType<typeof analyze>;
		try {
			probed = analyze(options);
		} finally {
			registry.splice(registry.indexOf(probe), 1);
		}
		expect(probed.failures).toEqual([{ stage: 'statement-walk', rule: 'probeStatements', message: 'statement visitor failed' }]);
		expect(probed.codes).toEqual(baseline.codes);
	});

	it('stops only the expression visitor that throws, once, and the walk goes on for the rest', () => {
		const { options } = project();
		const baseline = analyze(options);
		const probe: DiagnosticRuleEntry = {
			name: 'probeExpressions',
			procedureExpressions: () => () => { throw new Error('expression visitor failed'); },
		};
		registry.unshift(probe);
		let probed: ReturnType<typeof analyze>;
		try {
			probed = analyze(options);
		} finally {
			registry.splice(registry.indexOf(probe), 1);
		}
		expect(probed.failures).toEqual([{ stage: 'expression-walk', rule: 'probeExpressions', message: 'expression visitor failed' }]);
		expect(probed.codes).toEqual(baseline.codes);
	});

	it('reports the whole pass failing, which returns nothing', () => {
		const { options } = project();
		const failed = analyze({ ...options, parsedModule: {} as never });
		expect(failed.codes).toEqual([]);
		expect(failed.failures).toHaveLength(1);
		expect(failed.failures[0].stage).toBe('analysis');
	});

	it('carries a failure out of the extension s analysis, and through the worker, as data', () => {
		const { options } = project();
		const clean = analyzeVbaModuleSource({ source: CALLER, ...options });
		expect(clean.analysisFailures).toBeUndefined();

		const probe: DiagnosticRuleEntry = { name: 'probeRun', run: () => { throw new Error('run failed'); } };
		const inHost = withRule(probe, () => analyzeVbaModuleSource({ source: CALLER, ...options }));
		expect(inHost.analysisFailures).toEqual([
			expect.objectContaining({ stage: 'rule', rule: 'probeRun', message: 'run failed', stack: expect.stringContaining('run failed') }),
		]);
		// The same diagnostics as a clean run: only the failed rule is missing.
		expect(inHost.diagnostics).toEqual(clean.diagnostics);

		const worker = new AnalysisWorkerState();
		worker.handle({
			kind: 'seed', projectKey: 'wb', generation: 1,
			modules: [{ moduleName: 'Lib', source: LIB, type: 'standard' }, { moduleName: 'Caller', source: CALLER, type: 'standard' }],
		});
		const response = withRule(probe, () => worker.handle({
			kind: 'analyze', requestId: 1, docKey: 'doc', projectKey: 'wb', generation: 1,
			source: CALLER, moduleName: 'Caller', moduleType: 'standard',
		}));
		expect(response?.kind === 'result' && response.analysisFailures).toEqual([
			expect.objectContaining({ stage: 'rule', rule: 'probeRun', message: 'run failed' }),
		]);
	});

	it('leaves every project option out when the index fails partway, and says so', () => {
		const { index } = project();
		// The fourth of the index's answers: the three before it used to stay set.
		index.visibleTypeNames = () => { throw new Error('index failed'); };
		const options = projectAnalysisOptionsForModule(index, 'Caller');
		expect(options.knownProcedures).toBeUndefined();
		expect(options.knownIdentifiers).toBeUndefined();
		expect(options.projectContextFailure).toBeInstanceOf(Error);

		const result = analyzeVbaModuleSource({ source: CALLER, moduleName: 'Caller', moduleKind: 'standard', ...options });
		expect(result.analysisFailures).toEqual([
			expect.objectContaining({ stage: 'project-context', message: 'index failed' }),
		]);
	});

	it('logs a module the live project index leaves out', () => {
		const lines: string[] = [];
		setAnalysisFailureLog((line) => lines.push(line));
		try {
			const project = buildLiveVbaProjectIndex([
				{ moduleName: 'Lib', type: 'standard', source: LIB },
				{ moduleName: 'Broken', type: 'standard', source: undefined as never },
			]);
			expect(project.getModule('Lib')).toBeDefined();
			expect(project.getModule('Broken')).toBeUndefined();
			expect(lines).toHaveLength(1);
			expect(lines[0]).toMatch(/^XLIDE could not fully check Broken: it was left out of its project, so no other module sees its declarations: /);
		} finally {
			setAnalysisFailureLog(undefined);
		}
	});

	it('writes each failure to the output channel once per module, with its stack', () => {
		const lines: string[] = [];
		setAnalysisFailureLog((line) => lines.push(line));
		try {
			const failure = { stage: 'rule' as const, rule: 'probeRun', message: 'run failed', stack: 'Error: run failed\n    at probe' };
			logAnalysisFailures('Book.xlsm/Caller.bas', [failure]);
			logAnalysisFailures('Book.xlsm/Caller.bas', [failure]);
			logAnalysisFailures('Book.xlsm/Other.bas', [failure]);
			expect(lines).toEqual([
				'XLIDE could not fully check Book.xlsm/Caller.bas: rule probeRun did not run: run failed\nError: run failed\n    at probe',
				'XLIDE could not fully check Book.xlsm/Other.bas: rule probeRun did not run: run failed\nError: run failed\n    at probe',
			]);
		} finally {
			setAnalysisFailureLog(undefined);
		}
	});

	it('still never throws when the callback does', () => {
		const { options } = project();
		const diagnostics = withRule({ name: 'probeRun', run: () => { throw new Error('run failed'); } }, () =>
			analyzeModule(CALLER, { ...options, onInternalError: () => { throw new Error('host callback failed'); } }));
		expect(diagnostics.length).toBeGreaterThan(0);
	});
});
