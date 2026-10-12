// A #Const whose value reads a name with no value (issue #1347). Measured
// in 64-bit Excel 16.0 (2026-10-13), with no project arguments set: a module
// #Const replaces a built-in's value (#Const Win64 = 0 and #Const VBA7 = 0
// make #If Win64 and #If VBA7 False), and a name nothing defines is 0 in it
// (#Const Win64 = UNDEFINED_X makes #If Win64 False). When the project's
// arguments are not known, XLIDE reads such a name as unknown (issue #102),
// so the #Const is unknown too: it must not keep the built-in's default.

import { describe, expect, it } from 'vitest';
import {
	createConditionalActivityTracker,
	parseModule,
	type ConditionalCompilationEnvironment,
} from '../src/analyzer';
import { conditionalActivityAtOffset } from '../src/analyzer/conditional/conditionalCompilation';
import { documentOutlineSymbolsForSource } from '../src/vbaSymbolPresentation';

const module = (lines: string) =>
	`${lines}\nSub Main()\n#If Win64 Then\n    Debug.Print 1\n#Else\n    Debug.Print 2\n#End If\nEnd Sub\n`;

/** The activity of the #If arm and the #Else arm. */
function arms(source: string, env?: ConditionalCompilationEnvironment): string[] {
	const parsed = parseModule(source);
	const tracker = createConditionalActivityTracker(parsed, env);
	expect(tracker).toBeDefined();
	return ['Debug.Print 1', 'Debug.Print 2'].map((text) => {
		const at = source.indexOf(text);
		expect(at).toBeGreaterThanOrEqual(0);
		const activity = tracker!.activityForSpan({ start: at, end: at + 1 });
		// The offset replay shares the change.
		expect(conditionalActivityAtOffset(parsed, at, env)).toBe(activity);
		return activity;
	});
}

describe('a #Const that reads a name with no value', () => {
	it('leaves a built-in unknown when the project arguments are not known', () => {
		expect(arms(module('#Const Win64 = UNDEFINED_X'))).toEqual(['unknown', 'unknown']);
		expect(arms(module('#Const Win64 = MISSING + 1'))).toEqual(['unknown', 'unknown']);
	});

	it('is decided as before when they are', () => {
		expect(arms(module('#Const Win64 = UNDEFINED_X'), { projectConstants: {} })).toEqual(['inactive', 'active']);
		expect(arms(module('#Const Win64 = MISSING + 1'), { projectConstants: {} })).toEqual(['active', 'inactive']);
		expect(arms(module('#Const Win64 = MISSING + 1'), { projectConstants: { MISSING: -1 } })).toEqual(['inactive', 'active']);
	});

	it('still takes a built-in value a #Const gives, and the default when none does', () => {
		expect(arms(module('#Const Win64 = 0'))).toEqual(['inactive', 'active']);
		expect(arms(module("' no #Const"))).toEqual(['active', 'inactive']);
	});

	it('leaves the TWINBASIC default unknown too', () => {
		const source = '#Const TWINBASIC = UNDEFINED_X\n#If TWINBASIC Then\nDim a\n#Else\nDim b\n#End If\n';
		const tracker = createConditionalActivityTracker(parseModule(source))!;
		const at = source.indexOf('Dim b');
		expect(tracker.activityForSpan({ start: at, end: at + 1 })).toBe('unknown');
	});

	it('makes a #Const that reads an unknown one unknown in turn', () => {
		// VBA7 has a default, so this passes only if its unknown value is read.
		expect(arms(module('#Const VBA7 = UNDEFINED_X\n#Const Win64 = VBA7 + 1'))).toEqual(['unknown', 'unknown']);
	});

	it('keeps the earlier value for a #Const that fails for another reason', () => {
		// These have no value whatever the names are; that path is unchanged,
		// a malformed one reading an unknown name included.
		for (const value of ['1 / 0', 'UNDEFINED_X +', '(UNDEFINED_X', 'UNDEFINED_X 1', 'UNDEFINED_X + ,', 'UNDEFINED_X + Then', 'UNDEFINED_X + "bad']) {
			expect(arms(module(`#Const Win64 = ${value}`)), value).toEqual(['active', 'inactive']);
		}
	});

	it('counts only ordinary identifiers as names with no value', () => {
		// A keyword or a bracketed name is read as before.
		expect(arms(module('#Const Win64 = [UNDEFINED_X]'))).toEqual(['active', 'inactive']);
	});

	it('lists B in the Outline, which reads no project arguments', () => {
		// Both arms are unknown there, so both are listed.
		const source = '#Const Win64 = UNDEFINED_X\n#If Win64 Then\nSub A()\nEnd Sub\n#Else\nSub B()\nEnd Sub\n#End If\n';
		expect(documentOutlineSymbolsForSource('Module1', 'standard', source).map((symbol) => symbol.name)).toEqual(['A', 'B']);
	});
});
