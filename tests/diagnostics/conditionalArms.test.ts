// Rules that read the shape of a module, across the arms of a `#If` chain.
//
// github.com/WilliamSmithEdward/xlide_vscode/issues/58 started with duplicate
// declarations, but the same fault ran through every rule that scans a module
// linearly and asks "have I seen this already". Only one arm of a chain is ever
// compiled, so two things in different arms are alternatives; a rule that does
// not ask reports a conflict no build ever has.
//
// A chain XLIDE can decide is not the interesting case: the losing arms are
// dropped as inactive before the rules run, which is why `#If VBA7` never
// misfired. Every test here therefore runs the same source twice, once with a
// condition XLIDE cannot decide and once with `VBA7`, and expects both to be
// quiet. The controls that follow prove the rules still catch the real fault.

import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { analyzeProjectModule } from './helpers';

type Kind = 'standard' | 'class';

function countFor(source: string, code: string, kind: Kind): number {
	return analyzeVbaModuleSource({
		source,
		moduleName: 'Mod1',
		moduleType: kind,
		moduleKind: kind as never,
	}).diagnostics.filter((d) => d.code === code).length;
}

/** Runs `template` with FLAG undecidable and with FLAG = VBA7 (decidable). */
function bothWays(template: string, code: string, kind: Kind = 'standard'): number[] {
	const build = (flag: string) => template.replace(/FLAG/g, flag).split('\n').join('\r\n') + '\r\n';
	return [
		countFor(build('CUSTOM_FLAG'), code, kind),
		countFor(build('VBA7'), code, kind),
	];
}

function count(lines: string[], code: string, kind: Kind = 'standard'): number {
	return countFor(lines.join('\r\n') + '\r\n', code, kind);
}

describe('a rule must not pair things from different arms', () => {
	it('a declaration in one arm is not after a procedure in another', () => {
		expect(bothWays(`Option Explicit
#If FLAG Then
Public Sub T()
End Sub
#Else
Public X As Long
#End If`, 'module-declaration-after-procedure')).toEqual([0, 0]);
	});

	it('an Option in one arm is not after a declaration in another', () => {
		expect(bothWays(`#If FLAG Then
Public X As Long
#Else
Option Explicit
#End If
Sub T()
End Sub`, 'option-after-declaration')).toEqual([0, 0]);
	});

	it('an Implements in one arm is not after a procedure in another', () => {
		expect(bothWays(`Option Explicit
#If FLAG Then
Public Sub T()
End Sub
#Else
Implements IFoo
#End If`, 'implements-statement-placement', 'class')).toEqual([0, 0]);
	});

	it('an ElseIf in one arm does not follow an Else in another', () => {
		expect(bothWays(`Option Explicit
Sub T()
    Dim x As Boolean, y As Boolean
    If x Then
#If FLAG Then
    Else
        Debug.Print 1
#Else
    ElseIf y Then
        Debug.Print 2
#End If
    End If
End Sub`, 'else-branch-order')).toEqual([0, 0]);
	});

	// This one needs more than exclusivity: the `For` and the `Next` sit in two
	// SEPARATE chains, which are not exclusive at all. The pairing is only
	// meaningful when both sit under the same arms, so the rule asks that.
	it('a Next in one chain is not paired against a For in another', () => {
		expect(bothWays(`Option Explicit
Sub T()
    Dim i As Long, j As Long
#If FLAG Then
    For i = 1 To 2
#Else
    For j = 1 To 2
#End If
        Debug.Print 1
#If FLAG Then
    Next i
#Else
    Next j
#End If
End Sub`, 'next-variable-mismatch')).toEqual([0, 0]);
	});
});

describe('the same rules still catch the real fault', () => {
	it('reports a declaration that really follows a procedure', () => {
		expect(count([
			'Option Explicit', 'Public Sub T()', 'End Sub', 'Public X As Long',
		], 'module-declaration-after-procedure')).toBe(1);
	});

	it('reports both inside one arm, where every build compiles them together', () => {
		expect(count([
			'Option Explicit', '#If CUSTOM_FLAG Then', 'Public Sub T()', 'End Sub',
			'Public X As Long', '#End If',
		], 'module-declaration-after-procedure')).toBe(1);
	});

	it('reports an Option that really follows a declaration', () => {
		expect(count(['Public X As Long', 'Option Explicit'], 'option-after-declaration')).toBe(1);
	});

	it('reports a Next that really names the wrong variable', () => {
		expect(count([
			'Option Explicit', 'Sub T()', 'Dim i As Long, j As Long',
			'For i = 1 To 2', 'Debug.Print 1', 'Next j', 'End Sub',
		], 'next-variable-mismatch')).toBe(1);
	});

	it('reports a mismatched Next inside a single arm', () => {
		expect(count([
			'Option Explicit', 'Sub T()', 'Dim i As Long, j As Long',
			'#If CUSTOM_FLAG Then', 'For i = 1 To 2', 'Debug.Print 1', 'Next j',
			'#End If', 'End Sub',
		], 'next-variable-mismatch')).toBe(1);
	});

	it('reports an ElseIf that really follows an Else', () => {
		expect(count([
			'Option Explicit', 'Sub T()', 'Dim x As Boolean, y As Boolean', 'If x Then',
			'Else', 'Debug.Print 1', 'ElseIf y Then', 'Debug.Print 2', 'End If', 'End Sub',
		], 'else-branch-order')).toBe(1);
	});

	it('reports an Implements that really follows a procedure', () => {
		expect(count([
			'Option Explicit', 'Public Sub T()', 'End Sub', 'Implements IFoo',
		], 'implements-statement-placement', 'class')).toBe(1);
	});
});

// The structural block-balance engine is separate from the rule engine, and
// had the same fault in both directions. Its exemption tested "same kind and
// same label" rather than the arm, so it missed every opener that DIFFERS
// between arms, and wrongly merged two identical openers inside ONE arm.
describe('the structural engine reads arms, not labels', () => {
	function balance(lines: string[]): number {
		return analyzeVbaModuleSource({
			source: lines.join('\r\n') + '\r\n',
			moduleName: 'Mod1',
			moduleType: 'standard',
		}).diagnostics.filter(
			(d) => d.code === 'missing-block-closer' || d.code === 'unmatched-block-closer',
		).length;
	}

	it('accepts a block opened differently in each arm and closed once below', () => {
		expect(balance(['Option Explicit', 'Sub T()',
			'#If CUSTOM_FLAG Then', 'With Sheet1', '#Else', 'With Sheet2', '#End If',
			'.Range("A1").Value = 1', 'End With', 'End Sub'])).toBe(0);
	});

	it('accepts a differing If, For and Do the same way', () => {
		expect(balance(['Option Explicit', 'Sub T()', 'Dim x As Long',
			'#If CUSTOM_FLAG Then', 'If x > 0 Then', '#Else', 'If x >= 0 Then', '#End If',
			'Debug.Print 1', 'End If', 'End Sub'])).toBe(0);
		expect(balance(['Option Explicit', 'Sub T()', 'Dim i As Long',
			'#If CUSTOM_FLAG Then', 'For i = 1 To 10', '#Else', 'For i = 1 To 20', '#End If',
			'Debug.Print i', 'Next i', 'End Sub'])).toBe(0);
		expect(balance(['Option Explicit', 'Sub T()', 'Dim x As Long',
			'#If CUSTOM_FLAG Then', 'Do While x < 2', '#Else', 'Do Until x >= 2', '#End If',
			'x = x + 1', 'Loop', 'End Sub'])).toBe(0);
	});

	it('accepts three arms, and procedures whose NAMES differ between arms', () => {
		expect(balance(['Option Explicit', 'Sub T()',
			'#If A_FLAG Then', 'With Sheet1', '#ElseIf B_FLAG Then', 'With Sheet2',
			'#Else', 'With Sheet3', '#End If', '.Range("A1").Value = 1', 'End With',
			'End Sub'])).toBe(0);
		expect(balance(['Option Explicit',
			'#If CUSTOM_FLAG Then', 'Public Sub Alpha()', '#Else', 'Public Sub Beta()',
			'#End If', 'Debug.Print 1', 'End Sub'])).toBe(0);
	});

	it('still reports two openers inside ONE arm, which really do need two closers', () => {
		expect(balance(['Option Explicit',
			'#If CUSTOM_FLAG Then', 'Sub T()', 'Sub T()', '#End If',
			'Debug.Print 1', 'End Sub'])).toBe(1);
	});

	it('still reports a genuinely unclosed block, inside an arm or outside one', () => {
		expect(balance(['Option Explicit', 'Sub T()',
			'With Sheet1', '.Range("A1").Value = 1', 'End Sub'])).toBe(1);
		expect(balance(['Option Explicit', 'Sub T()',
			'#If CUSTOM_FLAG Then', 'With Sheet1', '.Range("A1").Value = 1', '#End If',
			'End Sub'])).toBe(1);
	});
});

// A procedure declared once per arm gives the module several signatures for one
// name, and arity checking used to switch off entirely for that name. XLIDE
// cannot say which arm a build compiles, but it can say that NO arm accepts the
// call, which is wrong under every build.
describe('argument-count with a procedure declared once per arm', () => {
	const perArm = (call: string): string => [
		'Option Explicit',
		'#If CUSTOM_FLAG Then', 'Public Sub T(ByVal a As Long)', 'End Sub',
		'#Else', 'Public Sub T(ByVal a As Long, ByVal b As Long)', 'End Sub', '#End If',
		'Public Sub Drive()', `    ${call}`, 'End Sub',
	].join('\r\n') + '\r\n';

	const arity = (source: string): number =>
		analyzeProjectModule(source, [{ moduleName: 'Mod1', source }], 'Mod1')
			.filter((d) => d.code === 'argument-count').length;

	it('reports a call no arm accepts', () => {
		expect(arity(perArm('T 1, 2, 3, 4, 5'))).toBe(1);
		expect(arity(perArm('T'))).toBe(1);
	});

	it('accepts a call any arm accepts', () => {
		expect(arity(perArm('T 1'))).toBe(0);
		expect(arity(perArm('T 1, 2'))).toBe(0);
	});
});
