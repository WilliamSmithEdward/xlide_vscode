import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { diagnosticMetadataForCode } from '../../src/analyzer';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, effectiveModuleKind, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../src/vbaProjectAnalysis';
import { oracleModuleInput } from '../helpers/oracleModuleInput';
import { resolveRuntimeFunction } from '../../src/analyzer/runtime/vbaRuntime';
import { buildModuleSymbols } from '../../src/analyzer/symbols/buildModuleSymbols';
import { parseModule } from '../../src/analyzer/parser/parseModule';
import { checkArgumentCount } from '../../src/analyzer/diagnostics/rules/callArity';
import { statementTokens } from '../../src/analyzer/diagnostics/walker';
import type { LeafStatementNode, ProcedureNode } from '../../src/analyzer/parser/nodes';

const evidence = JSON.parse(readFileSync('syntax_corpus/oracle/evidence/gap_contexts_qualified_20261019.json', 'utf8'));
const neighbors = JSON.parse(readFileSync('syntax_corpus/oracle/evidence/gap_contexts_comparison_neighbors_20261019.json', 'utf8'));
const receivers = JSON.parse(readFileSync('syntax_corpus/oracle/evidence/gap_contexts_receiver_shadows_20261019.json', 'utf8'));
const results = new Map<string, any>([...evidence.checkpointRecords, ...neighbors.checkpointRecords, ...receivers.checkpointRecords].map((record: any) => [record.result.caseId, record.result]));
function analyzeProject(modules: Parameters<typeof buildVbaProjectIndex>[0]) {
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const signatures = projectProcedureSignatures(project);
	return modules.flatMap(mod => analyzeVbaModuleSource({
		source: mod.source, moduleName: mod.moduleName, moduleKind: effectiveModuleKind(mod),
		...projectAnalysisOptionsForModule(project, mod.moduleName, signatures),
	}).diagnostics.filter(d => diagnosticMetadataForCode(d.code ?? '')?.diagnosticKind === 'compile-error')
		.map(d => ({ ...d, moduleName: mod.moduleName, source: mod.source })));
}
function compile(body: string, declarations = '', parameters = '') {
	return analyzeProject([{ moduleName: 'Module1', type: 'standard', source: `Option Explicit\n${declarations}\nSub Main(${parameters})\n${body}\nEnd Sub` }]);
}

describe('measured qualified runtime compile constraints', () => {
	it('retains all 512 measured verdicts and their host build', () => {
		expect(evidence.cases).toHaveLength(512);
		expect(evidence.checkpointRecords).toHaveLength(512);
		expect(evidence.hostBuild).toBe('16.0.20430.20182');
	});
	it.each([...evidence.cases, ...neighbors.cases, ...receivers.cases].map((c: any) => [c.id, c]))('%s', (_id: string, c: any) => {
		const result = results.get(c.id);
		expect(['accepted', 'rejected']).toContain(result.outcome);
		const diagnostics = analyzeProject(c.modules.map(oracleModuleInput));
		if (result.outcome === 'accepted') { expect(diagnostics).toEqual([]); return; }
		const message = result.message.replace(/^Compile error:\s*/, '').trim();
		expect(diagnostics.some(d => d.message.toLowerCase().includes(message.toLowerCase())), JSON.stringify(diagnostics)).toBe(true);
		// Whole-line selections after dialog dismissal do not prove a narrow range.
		const selection = result.selection;
		if (selection?.startLine === selection?.endLine && selection.startColumn > 1
			&& diagnostics.some(d => ['is-operator-non-object', 'typeof-invalid-expression', 'argument-count'].includes(d.code))) {
			const selected = selection.source.split(/\r?\n/)[selection.startLine - 1].slice(selection.startColumn - 1, selection.endColumn - 1);
			expect(diagnostics.some(d => d.moduleName === selection.moduleName && d.source.slice(d.span.start, d.span.end) === selected), selected + JSON.stringify(diagnostics)).toBe(true);
		}
		if (/builtin_[lu]bound_/.test(c.id)) {
			expect(diagnostics.filter(d => ['argument-count', 'is-operator-non-object', 'typeof-invalid-expression'].includes(d.code))).toEqual([]);
		}
	});
	it.each([
		['Abs', '1'], ['Int', '1'], ['Fix', '1'], ['InStr', '"a", "a"'], ['InStrB', '"a", "a"'],
		['Len', '"a"'], ['LenB', '"a"'], ['Sgn', '1'], ['StrComp', '"a", "a"'],
	])('keeps bare %s constraints and display metadata separate from qualified calls', (name, args) => {
		const runtime = resolveRuntimeFunction(name)!;
		expect(runtime.qualifiedReturns).toBe('Variant');
		expect(runtime.signature).toContain(`As ${runtime.returns}`);
		expect(compile(`Debug.Print ${name}(${args}) Is Nothing`).some(d => d.code === 'is-operator-non-object')).toBe(true);
		for (const expression of [`VBA.${name}(${args}) Is Nothing`, `TypeOf VBA.${name}(${args}) Is Collection`]) {
			expect(compile(`Debug.Print ${expression}`), expression).toEqual([]);
		}
	});
	it.each(['InStr', 'InStrB', 'StrComp'])('retains measured Variant/Null/Object argument neighbors for bare %s', name => {
		for (const args of ['v, \"a\"', 'Null, \"a\"', 'obj, \"a\"', '\"a\", v']) {
			for (const prefix of ['', 'VBA.']) {
				for (const expression of [`${prefix}${name}(${args}) Is Nothing`, `TypeOf ${prefix}${name}(${args}) Is Collection`]) {
					expect(compile(`Dim v As Variant\nDim obj As Object\nDebug.Print ${expression}`), expression).toEqual([]);
				}
			}
		}
	});
	it.each(['LBound', 'UBound'])('leaves nonexistent VBA.%s exclusively to member availability even with invalid arity', name => {
		for (const args of ['', '1, 2, 3, 4']) {
			const hits = compile(`Debug.Print VBA.${name}(${args}) Is Nothing`);
			expect(hits.some(d => d.message.includes('Method or data member not found'))).toBe(true);
			expect(hits.filter(d => ['argument-count', 'is-operator-non-object', 'typeof-invalid-expression'].includes(d.code))).toEqual([]);
		}
	});
	// These binding controls accompany the independently replayed exact
	// receiver measurements above; declared Object receivers remain late bound.
	it.each(['Strings', 'Other', 'Module1'])('respects declared value receiver %s before source module signatures', name => {
		for (const binding of ['parameter', 'local', 'module']) {
			const declaration = binding === 'module' ? `Dim ${name} As Object\n` : '';
			const parameter = binding === 'parameter' ? `ByVal ${name} As Object` : '';
			const local = binding === 'local' ? `Dim ${name} As Object\n` : '';
			const fn = 'Public Function Foo(x As Long) As Long\nEnd Function\n';
			const source = `Option Explicit\n${declaration}${name === 'Module1' ? fn : ''}Sub Main(${parameter})\n${local}Debug.Print ${name}.Foo() Is Nothing\nEnd Sub`;
			const modules = [{ moduleName: 'Module1', type: 'standard', source }];
			if (name !== 'Module1') { modules.push({ moduleName: name, type: 'standard', source: fn }); }
			expect(analyzeProject(modules), binding).toEqual([]);
		}
	});
	it.each(['Strings', 'Other', 'Module1'])('retains true source module %s call arity', name => {
		const fn = 'Public Function Foo(x As Long) As Long\nEnd Function\n';
		const modules = [{ moduleName: 'Module1', type: 'standard', source: `${name === 'Module1' ? fn : ''}Sub Main()\nDebug.Print ${name}.Foo()\nEnd Sub` }];
		if (name !== 'Module1') { modules.push({ moduleName: name, type: 'standard', source: fn }); }
		const hits = analyzeProject(modules).filter(d => d.code === 'argument-count');
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain('Argument not optional');
		expect(hits[0].source.slice(hits[0].span.start, hits[0].span.end)).toBe('.Foo');
	});
	it.each(['[Len]()', 'Len()'])('distinguishes bracketed designer binding from reserved bare syntax: %s', call => {
		const source = `Sub Main()\nDebug.Print ${call} Is Nothing\nEnd Sub`;
		const parsedModule = parseModule(source);
		const proc = parsedModule.members[0] as ProcedureNode;
		const symbols = buildModuleSymbols('Form1', 'userform', source, { parsedModule });
		const messages: string[] = [];
		checkArgumentCount(source, symbols, undefined, undefined, { parsedModule, implicitMembers: [{ name: 'Len', type: 'MSForms.CommandButton' }] }, (_key, message) => messages.push(message))(proc)!(proc.body[0] as LeafStatementNode);
		if (call === '[Len]()') { expect(messages).toEqual([]); }
		else { expect(messages).toHaveLength(1); expect(messages[0]).toContain('Syntax error'); }
	});
	it('indexes token offsets linearly as qualified calls increase', () => {
		const samples = [25, 50, 100].map(count => {
			const source = `Sub Main()\nDebug.Print ${Array.from({ length: count }, () => 'VBA.CLng(1)').join('; ')}\nEnd Sub`;
			const parsedModule = parseModule(source);
			const proc = parsedModule.members[0] as ProcedureNode;
			const stmt = proc.body[0] as LeafStatementNode;
			const symbols = buildModuleSymbols('Module1', 'standard', source, { parsedModule });
			const tokens = statementTokens(source, stmt.span);
			let reads = 0;
			const starts = tokens.map(token => token.start);
			for (const token of tokens) {
				const start = token.start;
				Object.defineProperty(token, 'start', { configurable: true, get() { reads++; return start; } });
			}
			try {
				checkArgumentCount(source, symbols, undefined, undefined, { parsedModule }, () => {})(proc)!(stmt);
				return reads;
			} finally {
				tokens.forEach((token, index) => Object.defineProperty(token, 'start', { configurable: true, writable: true, value: starts[index] }));
			}
		});
		expect(samples[0]).toBeGreaterThan(0);
		expect(samples[1]).toBeLessThan(samples[0] * 2.3);
		expect(samples[2]).toBeLessThan(samples[1] * 2.3);
	});
});
