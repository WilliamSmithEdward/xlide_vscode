import { describe, expect, it } from 'vitest';
import { procedureHasUnstructuredFlow } from '../src/analyzer/flow/procedureUnstructured';
import { analyzeModule } from '../src/analyzer/diagnostics/analyzeModule';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
import { knownFunctionResults } from '../src/analyzer/diagnostics/functionResults';
import { moduleTypes } from '../src/analyzer/diagnostics/typeFields';
import { defaultedStraightLine, functionResultFor, unreachableStatementsIn } from '../src/analyzer/diagnostics/typeInference';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { tokenize } from '../src/analyzer/lexer/tokenize';
import { parseModule } from '../src/analyzer/parser/parseModule';

const environment = (flag: boolean) => ({ compilerConstants: { FLAG: flag } });
const errors = (source: string, flag: boolean) => analyzeModule(source, { conditionalCompilation: environment(flag) }).filter(d => d.severity === 'error').map(d => d.code);

describe('analyzer fact cache lifetimes', () => {
	it('changes literal function results when the active compile branch changes on a reused parse', () => {
		const source = '#If FLAG Then\nFunction F() As Long\nF = 0\nEnd Function\n#Else\nFunction F() As Long\nF = 2\nEnd Function\n#End If';
		const module = parseModule(source);
		for (const flag of [true, false, true]) {
			const activity = createConditionalActivityTracker(module, environment(flag));
			const facts = knownFunctionResults(source, module, activity);
			expect(facts.get('f')).toMatchObject({ value: flag ? 0 : 2 });
			expect(knownFunctionResults(source, module, activity)).toBe(facts);
		}
	});

	it('changes UDT array bounds when conditional constants change', () => {
		const source = '#If FLAG Then\nConst N As Long = 1\n#Else\nConst N As Long = 3\n#End If\nType T\na(0 To N) As Long\nEnd Type';
		const module = parseModule(source);
		for (const flag of [true, false, true]) {
			const activity = createConditionalActivityTracker(module, environment(flag));
			const facts = moduleTypes(source, module, activity);
			expect(facts.get('t')?.get('a')?.dims?.[0].upper).toBe(flag ? 1 : 3);
			expect(moduleTypes(source, module, activity)).toBe(facts);
		}
	});

	it('updates caller diagnostics in both directions after a compiler flag change', () => {
		const source = 'Option Explicit\n#If FLAG Then\nFunction F() As Long\nF = 0\nEnd Function\n#Else\nFunction F() As Long\nF = 2\nEnd Function\n#End If\nSub Main()\nDebug.Print 10 / F()\nEnd Sub';
		for (const flag of [true, false, true]) {
			expect(errors(source, flag).includes('division-by-zero')).toBe(flag);
		}
	});

	it('updates returned-array bounds after Option Base changes on the same parse', () => {
		const source = 'Option Explicit\n#If FLAG Then\nOption Base 1\n#Else\nOption Base 0\n#End If\nFunction F() As Variant\nF = Array(1, 2)\nEnd Function\nSub Main()\nDebug.Print F()(0)\nEnd Sub';
		for (const flag of [true, false, true]) {
			expect(errors(source, flag).includes('array-subscript-out-of-bounds')).toBe(flag);
		}
	});

	it('updates initial constant values for reachability after changing compile branches', () => {
		const source = 'Option Explicit\n#If FLAG Then\nConst Limit As Long = 0\n#Else\nConst Limit As Long = 1\n#End If\nSub Main()\nDim n As Long\nIf Limit = 0 Then Exit Sub\nn = 1 / 0\nDebug.Print n\nEnd Sub';
		for (const flag of [false, true, false]) {
			expect(errors(source, flag).includes('division-by-zero')).toBe(!flag);
		}
	});
	it('keeps symbol-dependent walk starts, reachability and calls within their symbol snapshot', () => {
		const source = 'Const Limit As Long = 0\nFunction F(ByVal n As Long) As Long\nIf Limit = 0 Then Exit Function\nF = Limit + n\nEnd Function\nFunction G(ByVal n As Long) As Long\nG = Limit + n\nEnd Function';
		const module = parseModule(source);
		const proc = module.members.find(member => member.kind === 'Procedure');
		if (proc?.kind !== 'Procedure') { throw new Error('Missing procedure'); }
		const base = buildModuleSymbols('M', 'standard', source, { parsedModule: module });
		const changed = { ...base, root: { ...base.root, children: base.root.children?.map(symbol =>
			symbol.kind === 'constant' ? { ...symbol, defaultRaw: '1' } : symbol) } };
		const calledProc = module.members.find(member => member.kind === 'Procedure' && member.name === 'G');
		if (calledProc?.kind !== 'Procedure') { throw new Error('Missing callable procedure'); }
		const resultNode = proc.body.find(node => node.kind === 'Assignment');
		if (!resultNode) { throw new Error('Missing assignment'); }
		for (const [symbols, expected] of [[base, '0'], [changed, '2'], [base, '0']] as const) {
			const dead = unreachableStatementsIn(source, proc, symbols, undefined);
			expect(dead.has(resultNode)).toBe(expected === '0');
			expect(unreachableStatementsIn(source, proc, symbols, undefined)).toBe(dead);
			const call = functionResultFor(source, calledProc, symbols, undefined, [tokenize('1').filter(token => token.kind !== 'eof')]);
			expect(call?.map(token => token.rawText).join('')).toBe(expected === '0' ? '1' : '2');
			expect(functionResultFor(source, calledProc, symbols, undefined, [tokenize('1').filter(token => token.kind !== 'eof')])).toBe(call);
			const walk = defaultedStraightLine(source, proc, symbols, undefined);
			expect(walk.get(resultNode)?.get('limit')?.map(token => token.rawText).join('')).toBe(expected === '0' ? undefined : '1');
		}
	});
	it('updates unstructured flow when a conditional error handler becomes inactive', () => {
		const source = 'Sub Main()\n#If FLAG Then\nOn Error Resume Next\n#End If\nEnd Sub';
		const module = parseModule(source);
		const proc = module.members.find(member => member.kind === 'Procedure');
		if (proc?.kind !== 'Procedure') { throw new Error('Missing procedure'); }
		for (const flag of [true, false, true]) {
			const activity = createConditionalActivityTracker(module, environment(flag));
			expect(procedureHasUnstructuredFlow(source, proc, activity)).toBe(flag);
		}
	});
});
