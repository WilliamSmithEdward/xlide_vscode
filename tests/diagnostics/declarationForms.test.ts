// Diagnostics tests: declaration compile errors from issue #124. Each
// refused sample was measured in Excel 16.0 (build 20326, 2026-09-25) with
// the VBE's message; each accepted one compiles there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const MAIN = 'Function Main() As Variant\n    Main = 1\nEnd Function\n';

describe('property-accessor-signature-mismatch - value type (issue #124)', () => {
	it('flags a Let whose value type differs from the Get return type', () => {
		const src = 'Option Explicit\nPrivate mSize As Long\nPublic Property Get Size() As Long\n    Size = mSize\nEnd Property\nPublic Property Let Size(ByVal v As Integer)\n    mSize = v\nEnd Property\n';
		expectDiagnostic(src, analyzeModule(src, { moduleKind: 'class' }), 'property-accessor-signature-mismatch', { span: 'v', message: ['As Integer', 'returns Long'] });
	});

	it('stays quiet when the types agree, including both Variant', () => {
		const src = 'Option Explicit\nPublic Property Get Size() As Long\nEnd Property\nPublic Property Let Size(ByVal v As Long)\nEnd Property\nPublic Property Get Any()\nEnd Property\nPublic Property Let Any(v)\nEnd Property\n';
		expect(byCode(analyzeModule(src, { moduleKind: 'class' }), 'property-accessor-signature-mismatch')).toHaveLength(0);
	});
});

describe('duplicate-declaration and duplicate-procedure (issue #124)', () => {
	it('flags a local named after its own Function', () => {
		const src = 'Option Explicit\nFunction Main() As Variant\n    Dim Main As Long\n    Main = 1\nEnd Function\n';
		expectDiagnostic(src, analyzeModule(src), 'duplicate-declaration', { span: 'Main', message: 'Duplicate declaration' });
	});

	it('flags a module variable sharing a procedure name', () => {
		const src = 'Option Explicit\nPrivate Helper As Long\nPrivate Sub Helper()\nEnd Sub\n' + MAIN;
		expectDiagnostic(src, analyzeModule(src), 'duplicate-procedure', { span: 'Helper', message: 'Ambiguous name' });
	});
});

describe('withevents-declaration - type (issue #124)', () => {
	it('flags Object and Collection', () => {
		for (const [type, message] of [['Object', 'names none'], ['Collection', 'does not source']]) {
			const src = `Option Explicit\nPrivate WithEvents c As ${type}\n`;
			expectDiagnostic(src, analyzeModule(src, { moduleKind: 'class' }), 'withevents-declaration', { span: 'c', message });
		}
	});

	it('leaves a host class alone', () => {
		const src = 'Option Explicit\nPrivate WithEvents app As Application\n';
		expect(byCode(analyzeModule(src, { moduleKind: 'class' }), 'withevents-declaration')).toHaveLength(0);
	});
});

describe('invalid-option-statement - Option Private Module in an object module (issue #124)', () => {
	it('flags it in a class and allows it in a standard module', () => {
		const src = 'Option Private Module\n';
		expectDiagnostic(src, analyzeModule(src, { moduleKind: 'class' }), 'invalid-option-statement', { span: 'Module', message: 'not permitted' });
		expect(byCode(analyzeModule(src, { moduleKind: 'standard' }), 'invalid-option-statement')).toHaveLength(0);
	});
});

describe('parameter defaults, Enum values and array parameters (issue #124)', () => {
	it('flags a default outside the parameter type range as Overflow', () => {
		for (const [decl, span] of [['Optional ByVal i As Integer = 40000', '40000'], ['Optional ByVal b As Byte = 256', '256']]) {
			const src = `Option Explicit\nPrivate Sub F(${decl})\nEnd Sub\n` + MAIN;
			expectDiagnostic(src, analyzeModule(src), 'parameter-default-type-mismatch', { span, message: 'Overflow' });
		}
		const quiet = 'Option Explicit\nPrivate Sub F(Optional ByVal i As Integer = 32767, Optional ByVal b As Byte = 255, Optional ByVal d As Double = 1E+300)\nEnd Sub\n' + MAIN;
		expect(byCode(analyzeModule(quiet), 'parameter-default-type-mismatch')).toHaveLength(0);
	});

	it('flags an Enum member past the Long range', () => {
		const src = 'Option Explicit\nPrivate Enum E\n    eA = 3000000000#\n    eB = 2147483647\nEnd Enum\n' + MAIN;
		expectDiagnostic(src, analyzeModule(src), 'const-overflow', { span: 'eA', message: ['Enum member', 'Overflow'] });
	});

	it('flags a ByVal array parameter and an Optional array parameter', () => {
		const byVal = 'Option Explicit\nPrivate Sub F(ByVal a() As Long)\nEnd Sub\n' + MAIN;
		expectDiagnostic(byVal, analyzeModule(byVal), 'array-parameter-form', { span: 'a', message: 'must be ByRef' });
		const optional = 'Option Explicit\nPrivate Sub F(Optional a() As Long)\nEnd Sub\n' + MAIN;
		expectDiagnostic(optional, analyzeModule(optional), 'array-parameter-form', { span: 'a', message: 'Optional' });
		const quiet = 'Option Explicit\nPrivate Sub F(ByRef a() As Long, b() As Long, ParamArray c() As Variant)\nEnd Sub\n' + MAIN;
		expect(byCode(analyzeModule(quiet), 'array-parameter-form')).toHaveLength(0);
	});
});

describe('duplicate-deftype and bracketed-variable-name (issue #124)', () => {
	it('flags a letter given a default type twice', () => {
		const src = 'Option Explicit\nDefLng A-Z\nDefStr S\n' + MAIN;
		expectDiagnostic(src, analyzeModule(src), 'duplicate-deftype', { span: 'DefStr', message: "'S'" });
		const quiet = 'Option Explicit\nDefLng A-M\nDefStr N-Z\n' + MAIN;
		expect(byCode(analyzeModule(quiet), 'duplicate-deftype')).toHaveLength(0);
	});

	it('flags a bracketed variable name and allows a bracketed Enum member', () => {
		const src = 'Option Explicit\nFunction Main() As Variant\n    Dim [my var] As Long\n    Main = 1\nEnd Function\n';
		expectDiagnostic(src, analyzeModule(src), 'bracketed-variable-name', { span: '[my var]', message: 'Syntax error' });
		const quiet = 'Option Explicit\nPrivate Enum E\n    [Two Words] = 2\nEnd Enum\nFunction Main() As Variant\n    Main = E.[Two Words]\nEnd Function\n';
		expect(byCode(analyzeModule(quiet), 'bracketed-variable-name')).toHaveLength(0);
	});
});
