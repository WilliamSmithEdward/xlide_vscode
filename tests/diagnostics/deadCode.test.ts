// Diagnostics tests: the dead-code rules - unused-variable,
// variable-never-read, unused-procedure and unreachable-code.
//
// Each is a fact about the module text. The tests pin what is reported and,
// at greater length, what is left alone: a mention the rule cannot see
// through (ByRef, a string, a host-called handler) must keep it quiet.

import { describe, expect, it } from 'vitest';
import { analyzeModule, resolveDiagnosticCodeActions } from '../../src/analyzer';
import { byCode, expectDiagnostic, expectDiagnostics, spanText } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

function hits(src: string, code: string, opts: Parameters<typeof analyzeModule>[1] = {}) {
	return byCode(analyzeModule(src, opts), code);
}

function wrap(...body: string[]): string {
	return `Option Explicit\nSub T()\n${body.map((line) => `    ${line}`).join('\n')}\nEnd Sub\n`;
}

describe('analyzeModule - unused-variable', () => {
	it('flags a local nothing mentions, at its name, as information', () => {
		const src = wrap('Dim x As Long', 'Dim y As Long', 'y = 1', 'Debug.Print y');
		const found = expectDiagnostic(src, analyzeModule(src), 'unused-variable', {
			severity: 'information',
			span: 'x',
			message: "Local variable 'x' is declared but never used.",
		});
		expect(found.data?.removeDeclaration?.variableName).toBe('x');
	});

	it('offers to remove the whole line when the declaration stands alone on it', () => {
		const src = wrap('Dim x As Long', 'Debug.Print 1');
		const [found] = hits(src, 'unused-variable');
		const edit = found.data?.removeDeclaration?.edit;
		expect(edit).toBeDefined();
		expect(src.slice(edit!.span.start, edit!.span.end)).toBe('    Dim x As Long\n');
		expect(edit!.newText).toBe('');
		const actions = resolveDiagnosticCodeActions(src, { code: 'unused-variable', message: found.message, span: found.span, data: found.data });
		expect(actions.map((action) => action.title)).toContain("Remove unused declaration of 'x'");
	});

	it('removes only its own name from a Dim list', () => {
		const src = wrap('Dim a As Long, b As Long, c As Long', 'a = 1: c = a', 'Debug.Print c');
		const [found] = hits(src, 'unused-variable');
		expect(spanText(src, found)).toBe('b');
		const edit = found.data?.removeDeclaration?.edit;
		expect(src.slice(edit!.span.start, edit!.span.end)).toBe('b As Long, ');
		const last = wrap('Dim a As Long, b As Long', 'a = 1', 'Debug.Print a');
		const [lastFound] = hits(last, 'unused-variable');
		const lastEdit = lastFound.data?.removeDeclaration?.edit;
		expect(last.slice(lastEdit!.span.start, lastEdit!.span.end)).toBe(', b As Long');
	});

	it("removes a module variable's doc comment with it, and leaves a local's comment alone", () => {
		// Left behind, the comment documented the declaration that came next -
		// here a Sub, whose parameter the doc comment check then reported.
		const src = [
			'Option Explicit',
			"''' <summary>Rows read so far.</summary>",
			'Private m_rows As Collection',
			'Public Sub Go(ByVal sheet As Object)',
			"    ''' kept",
			'    Dim unused As Long',
			'    Debug.Print sheet.Name',
			'End Sub',
			'',
		].join('\n');
		const [moduleLevel, local] = hits(src, 'unused-variable');
		const edit = (found: typeof moduleLevel) => found.data!.removeDeclaration!.edit;
		expect(src.slice(edit(moduleLevel).span.start, edit(moduleLevel).span.end))
			.toBe("''' <summary>Rows read so far.</summary>\nPrivate m_rows As Collection\n");
		expect(src.slice(edit(local).span.start, edit(local).span.end)).toBe('    Dim unused As Long\n');
	});

	it('offers no removal when the declaration shares its line with a statement', () => {
		const src = wrap('Dim x As Long: Debug.Print 1');
		const [found] = hits(src, 'unused-variable');
		expect(found.data?.removeDeclaration).toBeUndefined();
	});

	it('flags module-private variables and constants, but not Public ones', () => {
		const src = [
			'Option Explicit',
			'Private m_count As Long',
			'Dim m_name As String',
			'Const LIMIT As Long = 10',
			'Private Const OTHER As Long = 11',
			'Public g_visible As Long',
			'Public Const SHARED As Long = 1',
			'Sub T()',
			'    Debug.Print OTHER',
			'End Sub',
		].join('\n');
		expectDiagnostics(src, analyzeModule(src), 'unused-variable', [
			{ span: 'm_count', message: "Module-level variable 'm_count' is declared but never used." },
			{ span: 'm_name' },
			{ span: 'LIMIT', message: "Constant 'LIMIT' is declared but never used." },
		]);
	});

	it('flags a Static local and a local Const', () => {
		const src = wrap('Static calls As Long', 'Const RATE As Double = 0.5', 'Debug.Print 1');
		expectDiagnostics(src, analyzeModule(src), 'unused-variable', [{ span: 'calls' }, { span: 'RATE' }]);
	});

	it('stays quiet for a local read anywhere, in any position', () => {
		const src = wrap(
			'Dim a As Long, b As Long, c As Object, d(1 To 3) As Long, e As Long, f As Long',
			'Helper a',
			'If b > 1 Then Debug.Print 1',
			'Set c = Nothing',
			'Debug.Print c.Name',
			'd(1) = 2',
			'For e = 1 To 3: Next',
			'Debug.Print f',
		);
		expect(hits(src, 'unused-variable')).toHaveLength(0);
	});

	it('stays quiet when the only mention is inside an inactive #If arm', () => {
		const src = wrap('Dim x As Long', '#If False Then', 'Debug.Print x', '#End If');
		expect(hits(src, 'unused-variable')).toHaveLength(0);
	});

	it('does not count a member name or a named argument as a mention', () => {
		const src = wrap('Dim Name As String', 'Debug.Print obj.Name', 'Foo Name:=1');
		expectDiagnostic(src, analyzeModule(src), 'unused-variable', { span: 'Name' });
	});

	it('does not let a shadowing local count as a use of the module variable', () => {
		const src = [
			'Option Explicit',
			'Private total As Long',
			'Sub A()',
			'    Dim total As Long',
			'    total = 1',
			'    Debug.Print total',
			'End Sub',
			'Sub B(total As Long)',
			'    Debug.Print total',
			'End Sub',
		].join('\n');
		expectDiagnostic(src, analyzeModule(src), 'unused-variable', { span: 'total' });
	});

	it('skips WithEvents variables and declarations in inactive #If arms', () => {
		const src = [
			'Option Explicit',
			'Private WithEvents app As Application',
			'#If False Then',
			'Private gone As Long',
			'#End If',
			'Sub T()',
			'End Sub',
		].join('\n');
		expect(hits(src, 'unused-variable', { moduleKind: 'class' })).toHaveLength(0);
	});

	it('skips a variable with a member attribute', () => {
		const src = [
			'Private m_default As Long',
			'Attribute m_default.VB_VarUserMemId = 0',
			'Sub T()',
			'End Sub',
		].join('\n');
		expect(hits(src, 'unused-variable', { moduleKind: 'class' })).toHaveLength(0);
	});
});

describe('analyzeModule - variable-never-read', () => {
	it('flags a local that is only ever assigned', () => {
		const src = wrap('Dim x As Long', 'x = 1', 'x = 2');
		expectDiagnostic(src, analyzeModule(src), 'variable-never-read', {
			severity: 'information',
			span: 'x',
			message: "Variable 'x' is assigned but its value is never read.",
		});
		expect(hits(src, 'unused-variable')).toHaveLength(0);
	});

	it('flags a Set-only object variable and a module variable written by every procedure', () => {
		const src = [
			'Option Explicit',
			'Private last As String',
			'Sub A()',
			'    Dim o As Object',
			'    Set o = CreateObject("Scripting.Dictionary")',
			'    last = "a"',
			'End Sub',
			'Sub B()',
			'    last = "b"',
			'End Sub',
		].join('\n');
		expectDiagnostics(src, analyzeModule(src), 'variable-never-read', [{ span: 'last' }, { span: 'o' }]);
	});

	it('treats x = x + 1, ReDim Preserve, a For counter and a ByRef pass as reads', () => {
		const src = wrap(
			'Dim a As Long, b() As Long, i As Long, c As Long, d As String',
			'a = a + 1',
			'ReDim b(1)',
			'ReDim Preserve b(2)',
			'For i = 1 To 3: Next',
			'Fill c',
			'Mid(d, 1, 1) = "x"',
		);
		expect(hits(src, 'variable-never-read')).toHaveLength(0);
	});

	it('treats an element write and a member write as reads of the variable', () => {
		const src = wrap('Dim arr(3) As Long, o As Object', 'arr(1) = 1', 'Set o = Nothing', 'o.Value = 2');
		expect(hits(src, 'variable-never-read')).toHaveLength(0);
	});

	it('never reports a constant as never read', () => {
		const src = wrap('Const A As Long = 1', 'Debug.Print 2');
		expect(hits(src, 'variable-never-read')).toHaveLength(0);
	});
});

describe('analyzeModule - unused-procedure', () => {
	it('flags a Private Sub, Function and Property nothing names', () => {
		const src = [
			'Option Explicit',
			'Private Sub Helper()',
			'End Sub',
			'Private Function Twice(n As Long) As Long',
			'    Twice = n * 2',
			'End Function',
			'Private Property Get Secret() As Long',
			'End Property',
			'Public Sub Main()',
			'End Sub',
		].join('\n');
		expectDiagnostics(src, analyzeModule(src), 'unused-procedure', [
			{ severity: 'information', span: 'Helper', message: "Private Sub 'Helper' is never called." },
			{ span: 'Twice', message: "Private Function 'Twice' is never called." },
			{ span: 'Secret', message: "Private Property 'Secret' is never used." },
		]);
	});

	it('stays quiet for a Private procedure the module names anywhere', () => {
		const src = [
			'Private Sub A()',
			'End Sub',
			'Private Function B() As Long',
			'End Function',
			'Private Sub C()',
			'End Sub',
			'Private Sub D()',
			'End Sub',
			'Public Sub Main()',
			'    A',
			'    Debug.Print B()',
			'    Application.OnTime Now, "C"',
			'    Dim p As LongPtr: p = AddressOf D',
			'End Sub',
		].join('\n');
		expect(hits(src, 'unused-procedure')).toHaveLength(0);
	});

	it('stays quiet when another module names the procedure in a string', () => {
		const target = 'Private Sub Refresh()\nEnd Sub\n';
		const caller = 'Sub Go()\n    Application.Run "Sheet1.Refresh"\nEnd Sub\n';
		const diags = analyzeProjectModule(target, [
			{ moduleName: 'Sheet1', source: target, type: 'document' },
			{ moduleName: 'Caller', source: caller },
		], 'Sheet1');
		expect(byCode(diags, 'unused-procedure')).toHaveLength(0);
		const alone = analyzeProjectModule(target, [
			{ moduleName: 'Sheet1', source: target, type: 'document' },
			{ moduleName: 'Caller', source: 'Sub Go()\nEnd Sub\n' },
		], 'Sheet1');
		expect(byCode(alone, 'unused-procedure')).toHaveLength(1);
	});

	it('never reports Public procedures', () => {
		const src = 'Public Sub Macro1()\nEnd Sub\nSub Macro2()\nEnd Sub\n';
		expect(hits(src, 'unused-procedure')).toHaveLength(0);
	});

	it('never reports event handlers, interface members or Auto_ macros', () => {
		const sheet = 'Private Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub\n';
		expect(hits(sheet, 'unused-procedure', { moduleKind: 'document' })).toHaveLength(0);
		const cls = 'Implements IFoo\nPrivate Sub Class_Initialize()\nEnd Sub\nPrivate Sub IFoo_Bar()\nEnd Sub\n';
		expect(hits(cls, 'unused-procedure', { moduleKind: 'class' })).toHaveLength(0);
		const form = 'Private Sub OkButton_Click()\nEnd Sub\nPrivate Sub UserForm_Initialize()\nEnd Sub\n';
		expect(hits(form, 'unused-procedure', { moduleKind: 'userform' })).toHaveLength(0);
		const auto = 'Private Sub Auto_Open()\nEnd Sub\n';
		expect(hits(auto, 'unused-procedure')).toHaveLength(0);
	});

	it('still reports an underscore-named Private Sub in a standard module', () => {
		const src = 'Private Sub Get_Data()\nEnd Sub\n';
		expectDiagnostic(src, analyzeModule(src), 'unused-procedure', { span: 'Get_Data' });
	});

	it('never reports a procedure carrying a member attribute', () => {
		const src = 'Private Sub Hot()\nAttribute Hot.VB_ProcData.VB_Invoke_Func = "h\\n14"\nEnd Sub\n';
		expect(hits(src, 'unused-procedure')).toHaveLength(0);
	});
});

describe('analyzeModule - unreachable-code', () => {
	it('flags the statements after Exit Sub up to the next label', () => {
		const src = [
			'Sub T()',
			'    On Error GoTo Fail',
			'    Debug.Print 1',
			'    Exit Sub',
			'    Debug.Print 2',
			'    Debug.Print 3',
			'Fail:',
			'    Resume Next',
			'End Sub',
		].join('\n');
		const found = expectDiagnostic(src, analyzeModule(src), 'unreachable-code', {
			severity: 'information',
			span: 'Debug.Print 2\n    Debug.Print 3',
			message: "Unreachable code after 'Exit Sub'.",
		});
		const edit = found.data?.removeUnreachableCode?.edit;
		expect(src.slice(edit!.span.start, edit!.span.end)).toBe('    Debug.Print 2\n    Debug.Print 3\n');
	});

	it('flags after GoTo, Resume, End, Return, Exit Do and Exit For', () => {
		const src = [
			'Sub T()',
			'    Do',
			'        Exit Do',
			'        a = 1',
			'    Loop',
			'    For i = 1 To 2',
			'        Exit For',
			'        b = 1',
			'    Next',
			'    If x Then',
			'        GoTo Done',
			'        c = 1',
			'    ElseIf y Then',
			'        End',
			'        d = 1',
			'    Else',
			'        Resume Next',
			'        e = 1',
			'    End If',
			'    GoSub Sub1',
			'Done:',
			'    Exit Sub',
			'Sub1:',
			'    Return',
			'    f = 1',
			'End Sub',
		].join('\n');
		expect(hits(src, 'unreachable-code').map((d) => spanText(src, d))).toEqual(['a = 1', 'b = 1', 'c = 1', 'd = 1', 'e = 1', 'f = 1']);
	});

	it('stays quiet across a Case arm, a line number and a #If', () => {
		const src = [
			'Sub T()',
			'    Select Case x',
			'        Case 1: Exit Sub',
			'        Case 2: Debug.Print 2',
			'    End Select',
			'    Exit Sub',
			'10  Debug.Print 3',
			'    Exit Sub',
			'#If DEBUGGING Then',
			'    Debug.Print 4',
			'#End If',
			'End Sub',
		].join('\n');
		expect(hits(src, 'unreachable-code')).toHaveLength(0);
	});

	it('does not treat a single-line If with Exit as unconditional', () => {
		const src = wrap('If x Then Exit Sub', 'Debug.Print 1');
		expect(hits(src, 'unreachable-code')).toHaveLength(0);
	});

	it('flags a whole block that follows an exit, unless a label inside it can be reached', () => {
		const dead = [
			'Sub T()',
			'    Exit Sub',
			'    If x Then',
			'        y = 1',
			'    End If',
			'End Sub',
		].join('\n');
		expectDiagnostic(dead, analyzeModule(dead), 'unreachable-code', { span: 'If x Then\n        y = 1\n    End If' });
		const reached = [
			'Sub T()',
			'    GoTo Inside',
			'    If x Then',
			'Inside:',
			'        y = 1',
			'    End If',
			'End Sub',
		].join('\n');
		expect(hits(reached, 'unreachable-code')).toHaveLength(0);
	});

	it('stays quiet for Exit Sub as the last statement of a block or a procedure', () => {
		const src = wrap('If x Then', '    Exit Sub', 'End If', 'Debug.Print 1', 'Exit Sub');
		expect(hits(src, 'unreachable-code')).toHaveLength(0);
	});
});
