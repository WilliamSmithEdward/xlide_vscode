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

	it('evaluates a relational operator in either order and split by a space (issue #87)', () => {
		expect(evaluateConditionalExpression('2 => 1')).toBe(true);
		expect(evaluateConditionalExpression('2 =< 1')).toBe(false);
		expect(evaluateConditionalExpression('2 >< 1')).toBe(true);
		expect(evaluateConditionalExpression('2 < > 2')).toBe(false);
		expect(evaluateConditionalExpression('1 > = 1')).toBe(true);
	});

	it('allows callers to override Win32 without deriving it from Win64', () => {
		expect(
			evaluateConditionalExpression('Win32 And Win64', {
				compilerConstants: { Win32: false, Win64: true },
			}),
		).toBe(false);
	});

	it('evaluates the operators as the VBE does (issue #192)', () => {
		// Each measured in 64-bit Excel 16.0: which #If branch compiles.
		const cases: Array<[string, boolean | number | string]> = [
			['Not 1', -2],
			['Not -1', 0],
			['Not 0', -1],
			['Not 1 = 2', true],
			['-0', -0],
			['1 And 2', 0],
			['1 Or 0 And 0', 1],
			['1 Eqv 1', -1],
			['0 Imp 0', -1],
			['3 Mod 2', 1],
			['3 \\ 2 = 1', true],
			['5 / 2 = 2.5', true],
			['2 ^ 3 = 8', true],
			['-2 ^ 2', -4],
			['(1 + 2) * 3 = 9', true],
			['&HFFFF = -1', true],
			['True = -1', true],
			['"a" < "b"', true],
			['"A" = "a"', true],
			['"x" & "y" = "xy"', true],
			['True And False', false],
		];
		for (const [expression, value] of cases) {
			expect(evaluateConditionalExpression(expression), expression).toBe(value);
		}
	});

	it('takes the branches 64-bit Office takes: Win32 is True, Win16 False (issue #192)', () => {
		const branch = (source: string): string => {
			const module = parseModule(source);
			const tracker = createConditionalActivityTracker(module, { projectConstants: {} });
			const at = source.indexOf('Main = "if"');
			return tracker?.activityForSpan({ start: at, end: at + 11 }) ?? 'none';
		};
		const wrap = (setup: string, condition: string): string =>
			`${setup}Function Main() As String\n#If ${condition} Then\n    Main = "if"\n#Else\n    Main = "else"\n#End If\nEnd Function\n`;
		expect(branch(wrap('', 'Win32'))).toBe('active');
		expect(branch(wrap('', 'Win16'))).toBe('inactive');
		expect(branch(wrap('#Const DEBUGGING = 1\n', 'Not DEBUGGING'))).toBe('active');
		// A #Const inside #If False still defines its constant.
		expect(branch(wrap('#If False Then\n#Const FAST = 1\n#End If\n', 'FAST'))).toBe('active');
		expect(branch(wrap('#Const A = 1\n#Const B = 1\n', 'A Xor B'))).toBe('inactive');
		expect(branch(wrap('', 'B + 1'))).toBe('active');
	});

	it('takes the branch Excel takes with Like, Empty, Null, Is and date literals (issue #208)', () => {
		// Each measured in 64-bit Excel 16.0, the condition alone in #If.
		const branch = (setup: string, condition: string): string => {
			const source = `${setup}Function Main() As String\n#If ${condition} Then\n    Main = "if"\n#Else\n    Main = "else"\n#End If\nEnd Function\n`;
			const tracker = createConditionalActivityTracker(parseModule(source), { projectConstants: {} });
			const at = source.indexOf('Main = "if"');
			const activity = tracker?.activityForSpan({ start: at, end: at + 11 });
			return activity === 'active' ? 'if' : activity === 'inactive' ? 'else' : 'unknown';
		};
		const measured: Array<[string, string, string]> = [
			['', '"abc" Like "a*"', 'if'],
			['', '"abc" Like "b*"', 'else'],
			['', '"b" Like "[a-c]"', 'if'],
			['#Const L = 3\n', 'L Like "3"', 'if'],
			['#Const L = 3\n', 'L Like "#"', 'if'],
			['', '"ABC" Like "a*"', 'if'],
			['', '"abc" Like "a?c"', 'if'],
			['', '"a1" Like "a#"', 'if'],
			['', '"b" Like "[!a]"', 'if'],
			['', '"a" Like "[A-C]"', 'if'],
			['', '3 Like 3', 'if'],
			['', 'True Like "True"', 'if'],
			['#Const D = 1\n', 'D <> Empty', 'if'],
			['#Const D = 0\n', 'D = Empty', 'if'],
			['', '0 = Empty', 'if'],
			['', 'Empty', 'else'],
			['', 'Not Empty', 'if'],
			['', 'Empty = ""', 'if'],
			['', 'Empty & "x" = "x"', 'if'],
			['', 'Empty + 1 = 1', 'if'],
			['#Const E = Empty\n', 'E & "x" = "x"', 'if'],
			['', 'UNDEFINED_X = Empty', 'if'],
			['', 'UNDEFINED_X & "x" = "x"', 'else'],
			['', 'Nothing Is Nothing', 'if'],
			['', 'Null Or True', 'if'],
			['', 'Null And False', 'else'],
			['', 'Null & "a" = "a"', 'if'],
			['', '#1/2/2000# > #1/1/2000#', 'if'],
			['', '#1/2/2000# - #1/1/2000# = 1', 'if'],
			['', '#1/1/2000# = 36526', 'if'],
			['', '#12:00:00 PM# = 0.5', 'if'],
			['', '#12:00:00 AM#', 'else'],
			['', '#1/1/2000#', 'if'],
		];
		for (const [setup, condition, expected] of measured) {
			expect(branch(setup, condition), condition).toBe(expected);
		}
		// Refused by the VBE, so neither branch is claimed.
		for (const condition of ['Null', 'Null = 1', 'Not Null', 'Empty Is Empty', 'Nothing', '0 = ""', '1 < "a"']) {
			expect(branch('', condition), condition).toBe('unknown');
		}
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
