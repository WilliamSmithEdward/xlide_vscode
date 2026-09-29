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

	it('flags a Variant Let beside a Long Get and a Long Let beside a Variant Get', () => {
		// Both refused in Excel 16.0 (issue #152, measured): a Let must match exactly.
		for (const [getType, letType] of [['Long', 'Variant'], ['Variant', 'Long']]) {
			const src = `Option Explicit\nPublic Property Get Item(ByVal i As Variant) As ${getType}\nEnd Property\nPublic Property Let Item(ByVal i As Variant, ByVal v As ${letType})\nEnd Property\n`;
			expectDiagnostic(src, analyzeModule(src, { moduleKind: 'class' }), 'property-accessor-signature-mismatch', { span: 'v', message: `As ${letType}` });
		}
	});

	it('never compares a Set value with the Get return type (issue #152)', () => {
		// Measured in Excel 16.0: every pair below compiles, a Long Get beside
		// an Object Set included, and the ROneCOne-shaped Item runs.
		const pairs = [['Variant', 'Object'], ['Variant', 'Collection'], ['Object', 'Collection'], ['Collection', 'Object'], ['Object', 'Variant'], ['Long', 'Object']];
		for (const [getType, setType] of pairs) {
			const src = `Option Explicit\nPublic Property Get Item(ByVal i As Variant) As ${getType}\nEnd Property\nPublic Property Set Item(ByVal i As Variant, ByVal v As ${setType})\nEnd Property\n`;
			expect(byCode(analyzeModule(src, { moduleKind: 'class' }), 'property-accessor-signature-mismatch')).toHaveLength(0);
		}
		const item =
			'Option Explicit\nPrivate mValue As Variant\n' +
			'Public Property Get Item(ByVal Index As Variant) As Variant\n    If IsObject(mValue) Then\n        Set Item = mValue\n    Else\n        Item = mValue\n    End If\nEnd Property\n' +
			'Public Property Let Item(ByVal Index As Variant, ByVal Value As Variant)\n    mValue = Value\nEnd Property\n' +
			'Public Property Set Item(ByVal Index As Variant, ByVal Value As Object)\n    Set mValue = Value\nEnd Property\n';
		expect(byCode(analyzeModule(item, { moduleKind: 'class' }), 'property-accessor-signature-mismatch')).toHaveLength(0);
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

describe('hex literals are signed by their width (issue #141)', () => {
	// Measured in Excel 16.0 (build 20326, 2026-09-26): the module compiles
	// and Main returns "-1073741824 -1 -32769"; &H10000 as an Integer default
	// is a compile error, Overflow.
	it('stays quiet on Win32 flag Enums and hex Optional defaults', () => {
		const src =
			'Option Explicit\n' +
			'Private Enum FileAccessFlags\n    GENERIC_READ = &H80000000\n    GENERIC_WRITE = &H40000000\n    ALL_BITS = &HFFFFFFFF\nEnd Enum\n' +
			'Private Function Mask(Optional ByVal m As Long = &HFFFFFFFF, Optional ByVal i As Integer = &H8000) As Double\n    Mask = CDbl(m) + i\nEnd Function\n' +
			'Function Main() As String\n    Main = (GENERIC_READ Or GENERIC_WRITE) & " " & ALL_BITS & " " & Mask()\nEnd Function\n';
		const diags = analyzeModule(src);
		expect(byCode(diags, 'const-overflow')).toHaveLength(0);
		expect(byCode(diags, 'parameter-default-type-mismatch')).toHaveLength(0);
	});

	it('still reports a five-digit hex default on an Integer', () => {
		const src = 'Option Explicit\nPrivate Function Mask(Optional ByVal i As Integer = &H10000) As Long\n    Mask = i\nEnd Function\n' + MAIN;
		expectDiagnostic(src, analyzeModule(src), 'parameter-default-type-mismatch', { span: '&H10000', message: 'Overflow' });
	});
});
