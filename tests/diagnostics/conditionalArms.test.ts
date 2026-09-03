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
