// Diagnostics tests: property index names, properties used in a form their
// procedures do not allow, Event declarations, and RaiseEvent arguments
// (issue #266). Every case was measured in 64-bit Excel 16.0 (build 20326,
// 2026-10-01) by compiling the whole project, or running it for the
// run-time cases.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { analyzeProjectModule } from './helpers';

function classErrors(cls: string, main = 'Main = 1'): Array<{ code: string; message: string }> {
	const mainSrc = `Option Explicit\nFunction Main() As Variant\n    Dim c As New Class1, o As Object\n    ${main}\nEnd Function\n`;
	const classSrc = `Option Explicit\n${cls}\n`;
	return [
		...analyzeProjectModule(mainSrc, [{ moduleName: 'Module1', source: mainSrc }, { moduleName: 'Class1', source: classSrc, type: 'class' }], 'Module1'),
		...analyzeProjectModule(classSrc, [{ moduleName: 'Module1', source: mainSrc }, { moduleName: 'Class1', source: classSrc, type: 'class' }], 'Class1', { moduleKind: 'class' }),
	].filter((diag) => diag.severity === 'error').map((diag) => ({ code: diag.code, message: diag.message }));
}

const get = (params: string, type = 'Long'): string => `Public Property Get P(${params}) As ${type}\nEnd Property`;
const letP = (params: string): string => `Public Property Let P(${params})\nEnd Property`;
const setP = (params: string): string => `Public Property Set P(${params})\nEnd Property`;

describe('property index parameters keep their names', () => {
	it.each([
		['a Let renaming the index', [get('ByVal i As Long'), letP('ByVal k As Long, ByVal v As Long')], "Index parameter 1 must keep its name: expected 'i', found 'k'"],
		['a Set renaming the index', [get('ByVal i As Long', 'Object'), setP('ByVal k As Long, ByVal v As Object')], "expected 'i', found 'k'"],
		['a Set renaming a Let\'s index', [letP('ByVal i As Long, ByVal v As Variant'), setP('ByVal k As Long, ByVal v As Object')], "Property Set 'P' argument list must match Property Let 'P'"],
		['two indexes swapped', [get('ByVal a As Long, ByVal b As Long'), letP('ByVal b As Long, ByVal a As Long, ByVal v As Long')], "Index parameter 1 must keep its name: expected 'a', found 'b'"],
		['the second index renamed', [get('ByVal a As Long, ByVal b As Long'), letP('ByVal a As Long, ByVal c As Long, ByVal v As Long')], "Index parameter 2 must keep its name"],
		['Optional indexes renamed', [get('Optional ByVal a As Long'), letP('Optional ByVal b As Long, ByVal v As Long')], "expected 'a', found 'b'"],
	])('reports %s', (_label, procs, message) => {
		const found = classErrors(procs.join('\n')).filter((diag) => diag.code === 'property-accessor-signature-mismatch');
		expect(found).toHaveLength(1);
		expect(found[0].message).toContain(message);
	});

	it('reports it in a standard module too', () => {
		const src = `Option Explicit\nFunction Main() As Variant\n    Main = 1\nEnd Function\n${get('ByVal i As Long')}\n${letP('ByVal k As Long, ByVal v As Long')}\n`;
		expect(analyzeModule(src).map((diag) => diag.code)).toContain('property-accessor-signature-mismatch');
	});

	it('stays quiet on a change of case, the value parameter\'s name, or the same names', () => {
		for (const procs of [
			[get('ByVal i As Long'), letP('ByVal I As Long, ByVal v As Long')],
			[get('ByVal i As Long'), letP('ByVal i As Long, ByVal anything As Long')],
			[get('Optional ByVal a As Long'), letP('Optional ByVal a As Long, ByVal v As Long')],
			[letP('ByVal i As Long, ByVal v As Variant'), setP('ByVal i As Long, ByVal v As Object')],
		]) {
			expect(classErrors(procs.join('\n')), procs.join(' / ')).toEqual([]);
		}
	});
});

describe('a ParamArray in a Property Let', () => {
	it('reports a ParamArray left to take the value', () => {
		const found = classErrors(letP('ParamArray v() As Variant'));
		expect(found.map((diag) => diag.code)).toEqual(['property-setter-missing-value']);
		expect(found[0].message).toContain('Argument not optional');
	});

	it('takes a value parameter after the ParamArray, and a Get with one', () => {
		expect(classErrors(letP('ParamArray v() As Variant, ByVal x As Long'))).toEqual([]);
		expect(classErrors(get('ParamArray v() As Variant', 'Variant'))).toEqual([]);
	});
});

describe('a property used in a form its procedures do not allow', () => {
	it.each([
		['a read with only a Let', letP('ByVal v As Long'), 'Main = c.P', "'c.P' has a Property Let and no Property Get, so it has no value to read"],
		['a Set read with only a Set', setP('ByVal v As Object'), 'Set o = c.P', "'c.P' has a Property Set and no Property Get, so it has no value to read"],
		['a Get called as a statement', get(''), 'c.P', "'c.P' is a property, and a statement cannot call one"],
		['a Let called with an argument', letP('ByVal v As Long'), 'c.P 5', "'c.P' is a property, and a statement cannot call one"],
	])('reports %s', (_label, cls, main, message) => {
		const found = classErrors(cls, main).filter((diag) => diag.code === 'invalid-property-use');
		expect(found).toHaveLength(1);
		expect(found[0].message).toContain(`${message}. This is a VBE compile error: Invalid use of property.`);
	});

	it('stays quiet on an assignment, a read with a Get, and a late-bound receiver', () => {
		expect(classErrors(letP('ByVal v As Long'), 'c.P = 5').map((diag) => diag.code)).not.toContain('invalid-property-use');
		expect(classErrors(get('').replace('End Property', '    P = 3\nEnd Property'), 'Main = c.P')).toEqual([]);
		expect(classErrors(`${get('ByVal i As Long')}\n${letP('ByVal i As Long, ByVal v As Long')}`, 'c.P(1) = 5').map((diag) => diag.code)).not.toContain('invalid-property-use');
		expect(classErrors(letP('ByVal i As Long, ByVal v As Long'), 'c.P(1) = 5')).toEqual([]);
		// A member of what the property returns, as real projects call one.
		const returnsCollection = 'Public Property Get Items(ByVal i As Long) As Collection\nEnd Property\nPublic Property Get Bag() As Collection\nEnd Property';
		expect(classErrors(returnsCollection, 'c.Items(1).Add 5').map((diag) => diag.code)).not.toContain('invalid-property-use');
		expect(classErrors(returnsCollection, 'c.Bag.Add 5').map((diag) => diag.code)).not.toContain('invalid-property-use');
		expect(classErrors(letP('ByVal v As Long'), 'Set o = New Class1: Main = o.P').map((diag) => diag.code)).not.toContain('invalid-property-use');
		expect(classErrors(letP('ByVal v As Long'), 'Dim v As Variant: Set v = New Class1: Main = v.P').map((diag) => diag.code)).not.toContain('invalid-property-use');
	});
});

describe('Event declarations', () => {
	it.each([
		['Public Event Done() As Long', 'invalid-proc-header', 'Expected: end of statement'],
		['Public Event Done(ByVal a() As Long)', 'event-parameter-form', 'Array argument must be ByRef'],
		['Private Event Done()', 'invalid-proc-header', 'Expected: Sub or Function or Property'],
		['Friend Event Done()', 'invalid-proc-header', 'Expected: Sub or Function or Property'],
		['Public Event Done()\nPublic Event Done()', 'duplicate-procedure', "Ambiguous name detected: 'Done'"],
		['Public Event Done()\nPublic Sub Fire()\n    Done\nEnd Sub', 'unknown-call', "Sub or Function not defined: 'Done'"],
		['Public Event Done(n As Long)\nPublic Sub Fire()\n    Call Done(1)\nEnd Sub', 'unknown-call', "Sub or Function not defined: 'Done'"],
	])('reports %j', (cls, code, message) => {
		const found = classErrors(cls).filter((diag) => diag.code === code);
		expect(found).toHaveLength(1);
		expect(found[0].message).toContain(message);
	});

	it('leaves a standard module\'s Event to the module-kind rule', () => {
		const src = 'Option Explicit\nPrivate Event Done()\n';
		expect(analyzeModule(src).map((diag) => diag.code)).not.toContain('invalid-proc-header');
	});

	it('stays quiet on a ByRef array, an Event beside a Sub of its name, and a Sub it may call', () => {
		for (const cls of [
			'Public Event Done(a() As Long)',
			'Public Event Done()\nPublic Sub Done()\nEnd Sub',
			'Public Event Done()\nPublic Sub Done()\nEnd Sub\nPublic Sub Fire()\n    Done\nEnd Sub',
		]) {
			expect(classErrors(cls), cls).toEqual([]);
		}
	});
});

describe('RaiseEvent arguments', () => {
	const fire = (event: string, ...body: string[]): string => `Public Event ${event}\nPublic Sub Fire()\n${body.map((line) => `    ${line}`).join('\n')}\nEnd Sub`;

	it.each([
		[fire('Done(n As Long)', 'RaiseEvent Done(n:=1)'), 'malformed-statement', "'n:=' names one. This is a VBE compile error: Syntax error."],
		[fire('Done(a As Long, b As Long)', 'RaiseEvent Done(1, b:=2)'), 'malformed-statement', "'b:='"],
		[fire('Done(n As Long)', 'Dim i As Integer', 'RaiseEvent Done(i)'), 'byref-argument-type-mismatch', "'i' is declared As Integer, but parameter 'n' of Event 'Done' is ByRef As Long. This is a VBE compile error: ByRef argument type mismatch."],
		[fire('Done(n As Long)', 'Dim s As String', 'RaiseEvent Done(s)'), 'byref-argument-type-mismatch', "'s' is declared As String"],
		[fire('Done(ByVal n As Long)', 'RaiseEvent Done("abc")'), 'argument-type-mismatch', `but got "abc", which is no number. This will raise Run-time error '13': Type mismatch.`],
		[fire('Done(ByVal n As Integer)', 'RaiseEvent Done(40000)'), 'argument-type-mismatch', "but got 40000, which does not fit. This will raise Run-time error '6': Overflow."],
	])('reports %j', (cls, code, message) => {
		const found = classErrors(cls).filter((diag) => diag.code === code);
		expect(found).toHaveLength(1);
		expect(found[0].message).toContain(message);
	});

	it('stays quiet on the type itself, a copy, a literal, and a ByVal conversion', () => {
		for (const cls of [
			fire('Done(n As Long)', 'Dim i As Long', 'RaiseEvent Done(i)'),
			fire('Done(n As Long)', 'Dim i As Integer', 'RaiseEvent Done((i))'),
			fire('Done(n As Long)', 'RaiseEvent Done(5)'),
			fire('Done(n As Long)', 'Dim i As Integer', 'RaiseEvent Done(i + 1)'),
			fire('Done(ByVal n As Long)', 'Dim i As Integer', 'RaiseEvent Done(i)'),
			fire('Done(ByVal n As Long)', 'RaiseEvent Done("5")'),
		]) {
			expect(classErrors(cls), cls).toEqual([]);
		}
	});
});
