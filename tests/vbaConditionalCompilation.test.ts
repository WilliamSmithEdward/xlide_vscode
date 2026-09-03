import { describe, expect, it } from 'vitest';
import {
	collectConditionalDirectives,
	conditionalActivityAtOffset,
	createConditionalActivityTracker,
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

	it('defaults VBA7 branch activity to modern VBA for analyzer callers', () => {
		const source =
			'#If VBA7 Then\n' +
			'Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)\n' +
			'#Else\n' +
			'Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)\n' +
			'#End If\n';
		const module = parseModule(source);
		expect(conditionalActivityAtOffset(module, source.indexOf('PtrSafe'))).toBe('active');
		expect(conditionalActivityAtOffset(module, source.lastIndexOf('Declare Sub'))).toBe(
			'inactive',
		);
	});

	it('defaults platform branch activity to modern Windows 64-bit Office', () => {
		const source =
			'#If Win64 Then\n' +
			'Dim platform As LongPtr\n' +
			'#ElseIf Win32 Then\n' +
			'Dim platform As Long\n' +
			'#ElseIf Mac Then\n' +
			'Dim platform As Variant\n' +
			'#End If\n';
		const module = parseModule(source);
		expect(conditionalActivityAtOffset(module, source.indexOf('LongPtr'))).toBe('active');
		expect(conditionalActivityAtOffset(module, source.indexOf('Long\n'))).toBe('inactive');
		expect(conditionalActivityAtOffset(module, source.indexOf('Variant'))).toBe('inactive');
	});

	it('allows callers to override platform defaults for branch activity', () => {
		const source =
			'#If Win64 Then\n' +
			'Dim platform As LongPtr\n' +
			'#ElseIf Win32 Then\n' +
			'Dim platform As Long\n' +
			'#ElseIf Mac Then\n' +
			'Dim platform As Variant\n' +
			'#End If\n';
		const module = parseModule(source);
		const env = { compilerConstants: { Win64: false, Win32: false, Mac: true } };
		expect(conditionalActivityAtOffset(module, source.indexOf('LongPtr'), env)).toBe('inactive');
		expect(conditionalActivityAtOffset(module, source.indexOf('Long\n'), env)).toBe('inactive');
		expect(conditionalActivityAtOffset(module, source.indexOf('Variant'), env)).toBe('active');
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

	it('compares a boolean #Const equal to its VBA numeric value (False = 0)', () => {
		// `#Const Windows = (Mac = 0)` must be True on Windows (Mac is False = 0),
		// and `#If Windows And (TWINBASIC = 0)` must be active.
		const source =
			'#Const Windows = (Mac = 0)\n#If Windows And (TWINBASIC = 0) Then\nDim onWindows As Long\n#Else\nDim elsewhere As Long\n#End If\n';
		const module = parseModule(source);
		expect(conditionalActivityAtOffset(module, source.indexOf('onWindows'))).toBe('active');
		expect(conditionalActivityAtOffset(module, source.indexOf('elsewhere'))).toBe('inactive');
	});

	it('treats a TWINBASIC branch as inactive and its #Else as active (VBA target)', () => {
		// TWINBASIC is a twinBASIC-only compiler constant, undefined (False) in
		// Excel VBA; modern libraries gate twinBASIC-only intrinsics behind it.
		const source =
			'#If Mac Then\nmemmove a, b, c\n#ElseIf TWINBASIC Then\nPutMemPtr addr, val\n#Else\nDim ok As Long\n#End If\n';
		const module = parseModule(source);
		expect(conditionalActivityAtOffset(module, source.indexOf('PutMemPtr'))).toBe('inactive');
		expect(conditionalActivityAtOffset(module, source.indexOf('Dim ok'))).toBe('active');
	});
});

describe('conditional compilation mutual exclusion', () => {
	// Which arm wins is a build-time decision, but that AT MOST ONE wins is
	// known even when none of the conditions can be evaluated. That is what
	// lets the duplicate-declaration rules keep quiet about a name declared
	// once per arm.
	function exclusive(source: string, first: string, second: string): boolean {
		const tracker = createConditionalActivityTracker(parseModule(source));
		expect(tracker, 'the module has directives').toBeDefined();
		const at = (needle: string) => {
			const start = source.indexOf(needle);
			expect(start, needle).toBeGreaterThanOrEqual(0);
			return { start, end: start + needle.length };
		};
		return tracker!.mutuallyExclusive(at(first), at(second));
	}

	it('separates the arms of one chain, however many there are', () => {
		const source =
			'#If A Then\nDim first As Long\n' +
			'#ElseIf B Then\nDim second As Long\n' +
			'#Else\nDim third As Long\n#End If\n';
		expect(exclusive(source, 'first', 'second')).toBe(true);
		expect(exclusive(source, 'second', 'third')).toBe(true);
		expect(exclusive(source, 'first', 'third')).toBe(true);
	});

	it('joins statements in the same arm, across a nested chain', () => {
		const source =
			'#If A Then\nDim first As Long\n#If B Then\n#End If\nDim second As Long\n#End If\n';
		expect(exclusive(source, 'first', 'second')).toBe(false);
	});

	it('separates an outer arm from a chain nested in the other arm', () => {
		const source =
			'#If A Then\nDim first As Long\n' +
			'#Else\n#If B Then\nDim second As Long\n#End If\n#End If\n';
		expect(exclusive(source, 'first', 'second')).toBe(true);
	});

	it('joins arms of two separate chains, which can both compile', () => {
		const source =
			'#If A Then\nDim first As Long\n#End If\n' +
			'#If B Then\nDim second As Long\n#End If\n';
		expect(exclusive(source, 'first', 'second')).toBe(false);
	});

	it('joins unconditional code to any arm', () => {
		const source = 'Dim first As Long\n#If A Then\nDim second As Long\n#End If\n';
		expect(exclusive(source, 'first', 'second')).toBe(false);
	});
});
