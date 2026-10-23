// Diagnostics tests: class attributes (issue #256). The original cases were
// measured in Excel 16.0 (build 20326, 2026-10-01) from workbooks pyOpenVBA
// wrote. A default member returning a Collection was measured through the VBE
// oracle in build 20430: in Excel each value form below, and in Word,
// PowerPoint and Access the Variant return and local, Count, the second Let,
// Debug.Print and a String target. The untyped Function return is not
// measured on its own.

import { describe, it, expect } from 'vitest';
import { byCode, expectDiagnostic } from '../../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

const ITEMS = 'Private mItems As New Collection\nPublic Sub Add(ByVal v As Variant)\n    mItems.Add v\nEnd Sub\n';

const CLASSES: Record<string, string> = {
	fieldDefault: 'Public Value As Variant\nAttribute Value.VB_VarUserMemId = 0\n',
	enumNoAttribute: `${ITEMS}Public Property Get NewEnum() As IUnknown\n    Set NewEnum = mItems.[_NewEnum]\nEnd Property\n`,
	enumNone: ITEMS,
	enumCollection: `${ITEMS}Public Property Get NewEnum() As Collection\nAttribute NewEnum.VB_UserMemId = -4\n    Set NewEnum = mItems\nEnd Property\n`,
	enumOk: `${ITEMS}Public Property Get NewEnum() As IUnknown\nAttribute NewEnum.VB_UserMemId = -4\n    Set NewEnum = mItems.[_NewEnum]\nEnd Property\n`,
	enumFunction: `${ITEMS}Public Function NewEnum() As IUnknown\nAttribute NewEnum.VB_UserMemId = -4\n    Set NewEnum = mItems.[_NewEnum]\nEnd Function\n`,
	noDefault: 'Public Value As Variant\n',
	defaultItem: 'Public Property Get Item(ByVal i As Long) As Variant\nAttribute Item.VB_UserMemId = 0\n    Item = i\nEnd Property\n',
	defaultItems: `${ITEMS}Public Property Get Items() As Collection\nAttribute Items.VB_UserMemId = 0\n    Set Items = mItems\nEnd Property\n`,
	defaultValue: 'Private m As Long\nPublic Property Get Value() As Long\nAttribute Value.VB_UserMemId = 0\n    Value = m\nEnd Property\n',
	defaultValueLet: 'Private m As Long\nPublic Property Get Value() As Long\nAttribute Value.VB_UserMemId = 0\n    Value = m\nEnd Property\nPublic Property Let Value(ByVal v As Long)\n    m = v\nEnd Property\n',
	defaultOptional: 'Public Property Get Item(Optional ByVal i As Long = 1) As Variant\nAttribute Item.VB_UserMemId = 0\n    Item = i\nEnd Property\n',
};

function analyze(cls: string, ...lines: string[]) {
	const src = `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	const project = [{ moduleName: 'Module1', source: src }, { moduleName: 'C1', type: 'class' as const, source: `Option Explicit\n${CLASSES[cls]}` }];
	return { src, diags: analyzeProjectModule(src, project, 'Module1') };
}

function errors(cls: string, ...lines: string[]): string[] {
	return analyze(cls, ...lines).diags.filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a Public field marked VB_VarUserMemId = 0', () => {
	it('is the default member, for a Let and a read', () => {
		expect(errors('fieldDefault', 'Dim c As New C1', 'c = 3', 'Main = c')).toEqual([]);
	});
});

describe('For Each over a class', () => {
	it.each([
		['enumNoAttribute', "no member marked VB_UserMemId = -4", "'438'"],
		['enumNone', "no member marked VB_UserMemId = -4", "'438'"],
		['enumCollection', 'returns a Collection', "'451'"],
	])('needs a -4 member that returns an object: %s', (cls, message, error) => {
		const { src, diags } = analyze(cls, 'Dim c As New C1, v As Variant', 'c.Add 2', 'For Each v In c', 'Next', 'Main = 1');
		expectDiagnostic(src, byCode(diags, 'object-default-value'), 'object-default-value', { span: 'c', message: [message, error] });
	});

	it.each(['enumOk', 'enumFunction'])('runs with one: %s', (cls) => {
		expect(errors(cls, 'Dim c As New C1, v As Variant', 'c.Add 2', 'For Each v In c', '    Main = v', 'Next')).toEqual([]);
	});
});

describe('a default member read the wrong way', () => {
	it.each([
		['noDefault', ['Dim c As New C1', 'Main = c(1)'], 'c', ['no default member to take an index', "'438'"]],
		['defaultItem', ['Dim c As New C1', 'Main = c'], 'c', ['takes an argument', "'449'"]],
	])('%s', (cls, lines, span, message) => {
		const { src, diags } = analyze(cls, ...lines);
		expectDiagnostic(src, byCode(diags, 'object-default-value'), 'object-default-value', { span, message });
	});

	describe.each(['excel', 'word', 'powerpoint', 'access'] as const)('Collection-returning defaults in %s', (host) => {
		// Let reads Items once and keeps its object.
		it.each([
			['Main = c'],
			['Dim v As Variant', 'v = c'],
			['Dim v', 'v = c'],
			['Dim a(1) As Variant', 'a(0) = c'],
			['c.Add 1', 'Dim v As Variant', 'v = c', 'If v.Count <> 1 Then Err.Raise 5'],
			['Take (c)'],
		])('keeps an object returned by a default member: %s', (...lines) => {
			const src = `Option Explicit\nFunction Main() As Variant\n    Dim c As New C1\n${lines.join('\n')}\nEnd Function\nSub Take(ByVal v As Variant)\nEnd Sub`;
			const project = [{ moduleName: 'Module1', source: src }, { moduleName: 'C1', type: 'class' as const, source: CLASSES.defaultItems }];
			expect(analyzeProjectModule(src, project, 'Module1', { host }).filter(diag => diag.severity === 'error')).toEqual([]);
		});

		it('also keeps the object in an untyped Function return', () => {
			const src = 'Function Main()\nDim c As New C1\nMain = c\nEnd Function';
			const project = [{ moduleName: 'Module1', source: src }, { moduleName: 'C1', type: 'class' as const, source: CLASSES.defaultItems }];
			expect(analyzeProjectModule(src, project, 'Module1', { host }).filter(diag => diag.severity === 'error')).toEqual([]);
		});

		it('does not claim 450 for Debug.Print (measured runtime 13)', () => {
			const src = 'Sub Main()\nDim c As New C1\nDebug.Print c\nEnd Sub';
			const project = [{ moduleName: 'Module1', source: src }, { moduleName: 'C1', type: 'class' as const, source: CLASSES.defaultItems }];
			expect(analyzeProjectModule(src, project, 'Module1', { host }).filter(diag => diag.severity === 'error')).toEqual([]);
		});

		it('does not attribute a second Let failure to the first Let in Get1', () => {
			const src = 'Function Get1() As Variant\nDim c As New C1\nGet1 = c\nEnd Function\nSub Main()\nDim v As Variant\nv = Get1()\nEnd Sub';
			const project = [{ moduleName: 'Module1', source: src }, { moduleName: 'C1', type: 'class' as const, source: CLASSES.defaultItems }];
			const lineStart = src.indexOf('Get1 = c');
			const lineEnd = src.indexOf('\n', lineStart);
			expect(byCode(analyzeProjectModule(src, project, 'Module1', { host }), 'object-default-value')
				.filter(diag => diag.span.start < lineEnd && diag.span.end > lineStart)).toEqual([]);
		});
	});

	it('reads Nothing from a variable never Set', () => {
		const { src, diags } = analyze('defaultValue', 'Dim c As C1', 'Main = c');
		expectDiagnostic(src, byCode(diags, 'object-variable-not-set'), 'object-variable-not-set', { span: 'c', message: "'91'" });
	});

	it('takes no Let through a Property Get alone', () => {
		const { src, diags } = analyze('defaultValue', 'Dim c As New C1', 'c = 5', 'Main = 1');
		expectDiagnostic(src, byCode(diags, 'readonly-member-assignment'), 'readonly-member-assignment', { span: 'c', message: 'Invalid use of property' });
	});

	it('stays quiet where it runs', () => {
		for (const [cls, ...lines] of [
			['defaultOptional', 'Dim c As New C1', 'Main = c'],
			['defaultItems', 'Dim c As New C1', 'c.Items.Add 4', 'Main = c(1)'],
			['defaultValueLet', 'Dim c As New C1', 'c = 5', 'Main = c'],
			['defaultItem', 'Dim c As New C1', 'Main = c(7)'],
			// An array of the class is indexed itself.
			['noDefault', 'Dim cs(1) As C1', 'Set cs(0) = New C1', 'cs(0).Value = 2', 'Main = cs(0).Value'],
		]) {
			expect(errors(cls, ...lines), lines.join(': ')).toEqual([]);
		}
	});
});
