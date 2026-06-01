import { describe, expect, it } from 'vitest';
import {
	collectConditionalDirectives,
	conditionalActivityAtOffset,
	evaluateConditionalExpression,
	indexConditionalCompilation,
} from '../src/analyzer/conditional/conditionalCompilation';
import { parseModule } from '../src/analyzer/parser/parseModule';

describe('conditional compilation index', () => {
	it('collects module and procedure directives in source order', () => {
		const module = parseModule(
			'#Const DEBUGGING = True\n' +
			'#If VBA7 Then\n' +
			'#End If\n' +
			'Sub T()\n' +
			'    #If DEBUGGING Then\n' +
			'    #End If\n' +
			'End Sub\n',
		);
		const directives = collectConditionalDirectives(module);
		expect(directives.map((hit) => hit.directive.directiveKind)).toEqual([
			'Const',
			'If',
			'EndIf',
			'If',
			'EndIf',
		]);
		expect(directives.map((hit) => hit.container.kind)).toEqual([
			'module',
			'module',
			'module',
			'procedure',
			'procedure',
		]);
	});

	it('indexes #Const definitions with high-confidence values', () => {
		const index = indexConditionalCompilation(
			parseModule('#Const DEBUGGING = VBA7 And Not Mac\n#Const LABEL = "dev"\n'),
			{ compilerConstants: { VBA7: true, Mac: false } },
		);
		expect(index.constants.map((constant) => [constant.name, constant.value])).toEqual([
			['DEBUGGING', true],
			['LABEL', 'dev'],
		]);
	});
});

describe('conditional compilation expression evaluation', () => {
	it('evaluates common compiler constant expressions', () => {
		const compilerConstants = { VBA7: true, Win64: true, Win32: true, Mac: false };
		expect(evaluateConditionalExpression('VBA7 And Win64', { compilerConstants })).toBe(true);
		expect(evaluateConditionalExpression('Mac Or Win64', { compilerConstants })).toBe(true);
		expect(evaluateConditionalExpression('Not Mac', { compilerConstants })).toBe(true);
		expect(evaluateConditionalExpression('Win64 = True', { compilerConstants })).toBe(true);
	});

	it('allows callers to override Win32 without deriving it from Win64', () => {
		expect(
			evaluateConditionalExpression('Win32 And Win64', {
				compilerConstants: { Win32: false, Win64: true },
			}),
		).toBe(false);
	});

	it('returns undefined for unknown expressions instead of guessing', () => {
		expect(evaluateConditionalExpression('VBA7')).toBeUndefined();
		expect(evaluateConditionalExpression('MissingConstant And VBA7')).toBeUndefined();
		expect(evaluateConditionalExpression('VBA7 + 1')).toBeUndefined();
	});
});

describe('conditional compilation branch activity', () => {
	it('marks mutually exclusive VBA7 branches active or inactive', () => {
		const source =
			'#If VBA7 Then\n' +
			'Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)\n' +
			'#Else\n' +
			'Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)\n' +
			'#End If\n';
		const module = parseModule(source);
		const env = { compilerConstants: { VBA7: true } };
		expect(conditionalActivityAtOffset(module, source.indexOf('PtrSafe'), env)).toBe('active');
		expect(conditionalActivityAtOffset(module, source.lastIndexOf('Declare Sub'), env)).toBe(
			'inactive',
		);
	});

	it('uses preceding active #Const values for later branches', () => {
		const source = '#Const DEBUGGING = True\n#If DEBUGGING Then\nDebug.Print "on"\n#End If\n';
		const module = parseModule(source);
		expect(conditionalActivityAtOffset(module, source.indexOf('Debug.Print'))).toBe('active');
	});

	it('keeps branches unknown when a condition cannot be proven', () => {
		const source = '#If SOME_HOST_FLAG Then\nDebug.Print "maybe"\n#End If\n';
		const module = parseModule(source);
		expect(conditionalActivityAtOffset(module, source.indexOf('Debug.Print'))).toBe('unknown');
	});
});
