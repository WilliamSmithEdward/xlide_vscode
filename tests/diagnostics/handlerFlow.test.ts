// Diagnostics tests: runtime errors control flow raises (issue #117). Each
// raising sample was measured in Excel 16.0 (build 20326, 2026-09-26); each
// quiet one runs there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

describe('handler-fall-through (issue #117)', () => {
	const CODE = 'handler-fall-through';

	it('flags a handler entered from above that re-raises Err.Number', () => {
		const src =
			'Option Explicit\nFunction Main() As Variant\n' +
			'    On Error GoTo Handler\n' +
			'    Main = 1\n' +
			'Handler:\n' +
			'    Err.Raise Err.Number\n' +
			'End Function\n';
		expectDiagnostic(src, analyzeModule(src), CODE, { severity: 'error', span: 'Handler', message: ["error '5'", 'Exit Function'] });
	});

	it('stays quiet with an Exit before the label, a handler that resumes, or a block above it', () => {
		const exits =
			'Option Explicit\nFunction Main() As Variant\n' +
			'    On Error GoTo Handler\n' +
			'    Main = 1\n' +
			'    Exit Function\n' +
			'Handler:\n' +
			'    Err.Raise Err.Number\n' +
			'End Function\n';
		const resumes =
			'Option Explicit\nFunction Main() As Variant\n' +
			'    On Error GoTo Handler\n' +
			'    Main = 1\n' +
			'Handler:\n' +
			'    Resume Next\n' +
			'End Function\n';
		const blockAbove =
			'Option Explicit\nFunction Main() As Variant\n' +
			'    On Error GoTo Handler\n' +
			'    If Main = 0 Then\n        Exit Function\n    End If\n' +
			'Handler:\n' +
			'    Err.Raise Err.Number\n' +
			'End Function\n';
		for (const src of [exits, resumes, blockAbove]) {
			expect(byCode(analyzeModule(src), CODE), src).toHaveLength(0);
		}
	});
});

describe('resume-without-error (issue #117)', () => {
	const CODE = 'resume-without-error';

	it('flags Resume in a procedure that installs no handler', () => {
		const src = 'Option Explicit\nFunction Main() As Variant\n    Main = 1\n    Resume Next\nEnd Function\n';
		expectDiagnostic(src, analyzeModule(src), CODE, { severity: 'error', span: 'Resume', message: "error '20'" });
	});

	it('stays quiet where On Error GoTo installs a handler', () => {
		const src =
			'Option Explicit\nFunction Main() As Variant\n' +
			'    On Error GoTo Handler\n' +
			'    Main = 1\n' +
			'    Exit Function\n' +
			'Handler:\n' +
			'    Resume Next\n' +
			'End Function\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});
});

describe('return-without-gosub (issue #117)', () => {
	const CODE = 'return-without-gosub';

	it('flags a GoSub target entered from the statement above it', () => {
		const src =
			'Option Explicit\nFunction Main() As Variant\n' +
			'    GoSub Work\n' +
			'Work:\n' +
			'    Main = Main + 1\n' +
			'    Return\n' +
			'End Function\n';
		expectDiagnostic(src, analyzeModule(src), CODE, { severity: 'error', span: 'Work', message: "error '3'" });
	});

	it('stays quiet when an Exit stands before the target', () => {
		const src =
			'Option Explicit\nFunction Main() As Variant\n' +
			'    GoSub Work\n' +
			'    Exit Function\n' +
			'Work:\n' +
			'    Main = Main + 1\n' +
			'    Return\n' +
			'End Function\n';
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});
});

describe('recursive-property-accessor (issue #117)', () => {
	const CODE = 'recursive-property-accessor';

	it('flags a Property Get that reads Me.<its own name>', () => {
		const src =
			'Option Explicit\nPrivate mName As String\n' +
			'Public Property Get Name() As String\n' +
			'    Name = Me.Name\n' +
			'End Property\n';
		expectDiagnostic(src, analyzeModule(src, { moduleKind: 'class' }), CODE, { severity: 'error', span: 'Me.Name', message: "error '28'" });
	});

	it('flags a Property Let that assigns its own name', () => {
		const src =
			'Option Explicit\nPrivate mAge As Long\n' +
			'Public Property Get Age() As Long\n    Age = mAge\nEnd Property\n' +
			'Public Property Let Age(ByVal v As Long)\n    Age = v\nEnd Property\n';
		expectDiagnostic(src, analyzeModule(src, { moduleKind: 'class' }), CODE, { span: 'Age', message: 'Property Let' });
	});

	it('stays quiet for the return variable in a Get and the backing field in a Let', () => {
		const src =
			'Option Explicit\nPrivate mName As String\n' +
			'Public Property Get Name() As String\n    Name = mName\n    Name = Name\nEnd Property\n' +
			'Public Property Let Name(ByVal v As String)\n    mName = v\nEnd Property\n';
		expect(byCode(analyzeModule(src, { moduleKind: 'class' }), CODE)).toHaveLength(0);
	});
});
