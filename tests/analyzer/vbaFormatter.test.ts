// The VBA formatter: indentation by block structure, keyword casing, the
// VBE's spacing, and the things it must never touch.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { formatVbaModule, tokenStreamDifference } from '../../src/analyzer/format/formatModule';

const SPACES = { tabSize: 4, insertSpaces: true };

function format(source: string, options = SPACES): string {
	const result = formatVbaModule(source, options);
	expect(result.refusal).toBeUndefined();
	return result.text as string;
}

function lines(...text: string[]): string {
	return text.join('\n');
}

function oracleSource(id: string): string {
	const corpus: { cases: { id: string; source: string }[] } = JSON.parse(
		readFileSync(join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
	);
	const found = corpus.cases.find((oracleCase) => oracleCase.id === id);
	if (!found) {
		throw new Error(`no oracle case ${id}`);
	}
	return found.source;
}

describe('formatVbaModule - indentation', () => {
	it('indents a procedure body one level and closes it at column 1', () => {
		expect(format(lines('Sub T()', 'x = 1', 'End Sub'))).toBe(lines('Sub T()', '    x = 1', 'End Sub'));
	});

	it('nests If, For, Do, While, With and Select bodies', () => {
		const src = lines(
			'Sub T()',
			'If a Then',
			'For i = 1 To 3',
			'Do',
			'While b',
			'With c',
			'Select Case d',
			'Case 1',
			'e = 2',
			'Case Else',
			'e = 3',
			'End Select',
			'End With',
			'Wend',
			'Loop',
			'Next',
			'ElseIf f Then',
			'g = 1',
			'Else',
			'g = 2',
			'End If',
			'End Sub',
		);
		expect(format(src)).toBe(lines(
			'Sub T()',
			'    If a Then',
			'        For i = 1 To 3',
			'            Do',
			'                While b',
			'                    With c',
			'                        Select Case d',
			'                            Case 1',
			'                                e = 2',
			'                            Case Else',
			'                                e = 3',
			'                        End Select',
			'                    End With',
			'                Wend',
			'            Loop',
			'        Next',
			'    ElseIf f Then',
			'        g = 1',
			'    Else',
			'        g = 2',
			'    End If',
			'End Sub',
		));
	});

	it('does not open a block for a single-line If, with or without a colon', () => {
		const src = lines(
			'Sub T()',
			'If a Then b = 1',
			'If a Then: b = 1',
			'If a Then b = 1 Else b = 2',
			'c = 3',
			'End Sub',
		);
		expect(format(src)).toBe(lines(
			'Sub T()',
			'    If a Then b = 1',
			'    If a Then: b = 1',
			'    If a Then b = 1 Else b = 2',
			'    c = 3',
			'End Sub',
		));
	});

	it('closes a block written on a one-line If inside that If, not the block around it (issue #128)', () => {
		const src = lines(
			'Sub A()',
			'With c',
			'If a Then With d: .Add 1: End With',
			'.Add 2',
			'If a Then .Add 3 Else For i = 1 To 2: .Add i: Next',
			'.Add 4',
			'10 If a Then With d: .Add 5: End With',
			'.Add 6',
			'End With',
			'End Sub',
		);
		expect(format(src)).toBe(lines(
			'Sub A()',
			'    With c',
			'        If a Then With d: .Add 1: End With',
			'        .Add 2',
			'        If a Then .Add 3 Else For i = 1 To 2: .Add i: Next',
			'        .Add 4',
			'10 If a Then With d: .Add 5: End With',
			'        .Add 6',
			'    End With',
			'End Sub',
		));
	});

	it('reads an Else after a colon as the one-line If\'s, not an outer block If\'s', () => {
		// VBE oracle if_single_line_else_after_colon_*: all accepted.
		const src = lines(
			'Sub T()',
			'If outer Then',
			'With c',
			'If a Then b = 1: Else c = 2',
			'.Add 1',
			'10 If a Then b = 1: Else: c = 2',
			'.Add 2',
			'If a Then If x Then c = 1: Else d = 2: Else e = 3',
			'.Add 3',
			'If a Then With d: End With: Else With e: End With',
			'.Add 4',
			'End With',
			'Else',
			'z = 1',
			'End If',
			'End Sub',
		);
		expect(format(src)).toBe(lines(
			'Sub T()',
			'    If outer Then',
			'        With c',
			'            If a Then b = 1: Else c = 2',
			'            .Add 1',
			'10 If a Then b = 1: Else: c = 2',
			'            .Add 2',
			'            If a Then If x Then c = 1: Else d = 2: Else e = 3',
			'            .Add 3',
			'            If a Then With d: End With: Else With e: End With',
			'            .Add 4',
			'        End With',
			'    Else',
			'        z = 1',
			'    End If',
			'End Sub',
		));
	});

	it('does not read a member named Else or Then as the one-line If\'s keyword', () => {
		const src = lines(
			'Sub T()',
			'With c',
			'If a Then With d.Else: End With',
			'.Add 1',
			'If a Then With d!Then: End With',
			'.Add 2',
			'If a Then o.Else = 1: Else b = 1',
			'.Add 3',
			'If x! Then With d: End With',
			'.Add 4',
			'If a Then Set x = d.Then',
			'.Add 5',
			'End With',
			'End Sub',
		);
		expect(format(src)).toBe(lines(
			'Sub T()',
			'    With c',
			'        If a Then With d.Else: End With',
			'        .Add 1',
			'        If a Then With d!Then: End With',
			'        .Add 2',
			'        If a Then o.Else = 1: Else b = 1',
			'        .Add 3',
			'        If x! Then With d: End With',
			'        .Add 4',
			'        If a Then Set x = d.Then',
			'        .Add 5',
			'    End With',
			'End Sub',
		));
	});

	it('does not open a block for If x Then: with nothing after the colon (issue #84)', () => {
		// The line after it runs unconditionally, and an End If after it does
		// not compile.
		const src = lines('Sub T()', 'If x Then:', 'x = False', '10 If x Then:', 'Debug.Print x', 'End Sub');
		expect(format(src)).toBe(lines('Sub T()', '    If x Then:', '    x = False', '10 If x Then:', '    Debug.Print x', 'End Sub'));
	});

	it('keeps a block If whose header ends in a comment', () => {
		const src = lines('Sub T()', "If a Then ' why", 'b = 1', 'End If', 'End Sub');
		expect(format(src)).toBe(lines('Sub T()', "    If a Then ' why", '        b = 1', '    End If', 'End Sub'));
	});

	it('closes a block If at the one-word EndIf, which the VBE reads as End If (issue #88)', () => {
		const src = lines('Sub T()', 'If a Then', 'Exit Sub', 'EndIf', 'x = 2', 'End Sub');
		expect(format(src)).toBe(lines('Sub T()', '    If a Then', '        Exit Sub', '    EndIf', '    x = 2', 'End Sub'));
	});

	it('closes one loop per name in Next i, j', () => {
		const src = lines('Sub T()', 'For i = 1 To 2', 'For j = 1 To 2', 'x = 1', 'Next j, i', 'y = 2', 'End Sub');
		expect(format(src)).toBe(lines(
			'Sub T()',
			'    For i = 1 To 2',
			'        For j = 1 To 2',
			'            x = 1',
			'    Next j, i',
			'    y = 2',
			'End Sub',
		));
	});

	it('handles Loop While / Loop Until and Do While / Do Until', () => {
		const src = lines('Sub T()', 'Do While a', 'b = 1', 'Loop', 'Do', 'c = 1', 'Loop Until d', 'End Sub');
		expect(format(src)).toBe(lines(
			'Sub T()',
			'    Do While a',
			'        b = 1',
			'    Loop',
			'    Do',
			'        c = 1',
			'    Loop Until d',
			'End Sub',
		));
	});

	it('indents Type and Enum bodies and treats a field named Type as a field', () => {
		const src = lines(
			'Private Type Rec',
			'Name As String',
			'Type As Long',
			'End Type',
			'Public Enum Color',
			'Red',
			'Green = 2',
			'End Enum',
		);
		expect(format(src)).toBe(lines(
			'Private Type Rec',
			'    Name As String',
			'    Type As Long',
			'End Type',
			'Public Enum Color',
			'    Red',
			'    Green = 2',
			'End Enum',
		));
	});

	it('treats a Type field named after a block keyword as a field', () => {
		const src = lines(
			'Private Type Rec',
			'For As Long',
			'Sub As Long',
			'Next As Long',
			'End As Long',
			'Attribute As Long',
			'With As New Collection',
			'While _',
			'As Long',
			'End Type',
			'Dim x As Long',
		);
		expect(format(src)).toBe(lines(
			'Private Type Rec',
			'    For As Long',
			'    Sub As Long',
			'    Next As Long',
			'    End As Long',
			'    Attribute As Long',
			'    With As New Collection',
			'    While _',
			'        As Long',
			'End Type',
			'Dim x As Long',
		));
	});

	it('formats the VBE oracle\'s keyword-named Type fields as fields', () => {
		const fields = oracleSource('reserved_member_name_type_fields_compile');
		expect(format(fields)).toBe(fields);
		// A #If inside the Type is still a directive, its arms one level in.
		expect(format(oracleSource('reserved_member_name_type_field_conditional_compile'))).toBe(lines(
			'Private Type Rec',
			'    #If Win64 Then',
			'        Next As LongLong',
			'    #Else',
			'        Next As Long',
			'    #End If',
			"    Loop As Long ' trailing comment",
			'End Type',
			'',
			'Public Sub XlideOracleEntry()',
			'End Sub',
			'',
		));
	});

	it('keeps labels and line numbers at column 1', () => {
		const src = lines(
			'Sub T()',
			'On Error GoTo Fail',
			'x = 1',
			'Exit Sub',
			'    Fail:',
			'Resume Next',
			'   10 y = 2',
			'Done: z = 3',
			'End Sub',
		);
		expect(format(src)).toBe(lines(
			'Sub T()',
			'    On Error GoTo Fail',
			'    x = 1',
			'    Exit Sub',
			'Fail:',
			'    Resume Next',
			'10 y = 2',
			'Done: z = 3',
			'End Sub',
		));
	});

	it('indents #If arms and keeps a module-level #If across procedures', () => {
		const src = lines(
			'#If Win64 Then',
			'Private Declare PtrSafe Sub A Lib "k" ()',
			'#Else',
			'Private Declare Sub A Lib "k" ()',
			'#End If',
			'Sub T()',
			'#If DEBUGGING Then',
			'Debug.Print 1',
			'#End If',
			'End Sub',
		);
		expect(format(src)).toBe(lines(
			'#If Win64 Then',
			'    Private Declare PtrSafe Sub A Lib "k" ()',
			'#Else',
			'    Private Declare Sub A Lib "k" ()',
			'#End If',
			'Sub T()',
			'    #If DEBUGGING Then',
			'        Debug.Print 1',
			'    #End If',
			'End Sub',
		));
	});

	it('keeps a procedure inside a module-level #If one level in', () => {
		const src = lines('#If A Then', 'Sub T()', 'x = 1', 'End Sub', '#End If');
		expect(format(src)).toBe(lines('#If A Then', '    Sub T()', '        x = 1', '    End Sub', '#End If'));
	});

	it('indents the body of a procedure whose header is written in each arm of a #If', () => {
		const src = lines(
			'#If VBA7 Then',
			'Public Sub Foo(ByVal p As LongPtr)',
			'#Else',
			'Public Sub Foo(ByVal p As Long)',
			'#End If',
			'x = 1',
			'End Sub',
			'Sub Bar()',
			'#If VBA7 Then',
			'y = 1',
			'#End If',
			'y = 2',
			'End Sub',
		);
		expect(format(src)).toBe(lines(
			'#If VBA7 Then',
			'    Public Sub Foo(ByVal p As LongPtr)',
			'#Else',
			'    Public Sub Foo(ByVal p As Long)',
			'#End If',
			'    x = 1',
			'End Sub',
			'Sub Bar()',
			'    #If VBA7 Then',
			'        y = 1',
			'    #End If',
			'    y = 2',
			'End Sub',
		));
	});

	it('carries a procedure opened in an earlier #If arm past #End If, unless an arm closed it', () => {
		const src = lines(
			'#If A Then',
			'Sub T()',
			'#Else',
			"' no header in this arm",
			'#End If',
			'x = 1',
			'End Sub',
			'#If A Then',
			'Sub U()',
			'#Else',
			'Sub U()',
			'y = 1',
			'End Sub',
			'#End If',
			'Dim z As Long',
		);
		expect(format(src)).toBe(lines(
			'#If A Then',
			'    Sub T()',
			'#Else',
			"    ' no header in this arm",
			'#End If',
			'    x = 1',
			'End Sub',
			'#If A Then',
			'    Sub U()',
			'#Else',
			'    Sub U()',
			'        y = 1',
			'    End Sub',
			'#End If',
			'Dim z As Long',
		));
	});

	it('indents comment lines and blank lines at the block level', () => {
		const src = lines('Sub T()', "' about x", '', 'x = 1', 'End Sub', '', "' module comment");
		expect(format(src)).toBe(lines('Sub T()', "    ' about x", '    ', '    x = 1', 'End Sub', '', "' module comment"));
	});

	it('gives a flat continuation line one level and keeps an aligned one aligned', () => {
		const src = lines(
			'Sub T()',
			'Call Foo(a, _',
			'b)',
			'Call Foo(a, _',
			'         b)',
			'End Sub',
		);
		expect(format(src)).toBe(lines(
			'Sub T()',
			'    Call Foo(a, _',
			'        b)',
			'    Call Foo(a, _',
			'             b)',
			'End Sub',
		));
	});

	it('does not read a continued If header as single-line', () => {
		const src = lines('Sub T()', 'If a And _', 'b Then', 'c = 1', 'End If', 'End Sub');
		expect(format(src)).toBe(lines('Sub T()', '    If a And _', '        b Then', '        c = 1', '    End If', 'End Sub'));
	});

	it('nets out openers and closers on one line', () => {
		const src = lines('Sub T()', 'Do: x = x + 1: Loop Until x > 3', 'With o: .a = 1: End With', 'y = 1', 'End Sub');
		expect(format(src)).toBe(lines(
			'Sub T()',
			'    Do: x = x + 1: Loop Until x > 3',
			'    With o: .a = 1: End With',
			'    y = 1',
			'End Sub',
		));
	});

	it('resets to module level at the next procedure header after an unclosed block', () => {
		const src = lines('Sub A()', 'If x Then', 'y = 1', 'End Sub', 'Sub B()', 'z = 1', 'End Sub');
		expect(format(src)).toBe(lines('Sub A()', '    If x Then', '        y = 1', 'End Sub', 'Sub B()', '    z = 1', 'End Sub'));
	});

	it('leaves a stray closer where the line is and does not go negative', () => {
		const src = lines('End If', 'Sub T()', 'x = 1', 'End Sub');
		expect(format(src)).toBe(lines('End If', 'Sub T()', '    x = 1', 'End Sub'));
	});

	it('indents Property procedures and Declare lines like their peers', () => {
		const src = lines(
			'Private Declare PtrSafe Function Tick Lib "kernel32" Alias "GetTickCount" () As Long',
			'Public Property Get Name() As String',
			'Name = m_name',
			'End Property',
			'Public Property Let Name(ByVal value As String)',
			'm_name = value',
			'End Property',
		);
		expect(format(src)).toBe(lines(
			'Private Declare PtrSafe Function Tick Lib "kernel32" Alias "GetTickCount" () As Long',
			'Public Property Get Name() As String',
			'    Name = m_name',
			'End Property',
			'Public Property Let Name(ByVal value As String)',
			'    m_name = value',
			'End Property',
		));
	});

	it('indents with tabs when asked', () => {
		expect(format(lines('Sub T()', 'If a Then', 'b = 1', 'End If', 'End Sub'), { tabSize: 4, insertSpaces: false }))
			.toBe(lines('Sub T()', '\tIf a Then', '\t\tb = 1', '\tEnd If', 'End Sub'));
	});

	it('reads existing tab indentation at the tab width when placing continuations', () => {
		const src = lines('Sub T()', '\tCall Foo(a, _', '\t\t\tb)', 'End Sub');
		expect(format(src)).toBe(lines('Sub T()', '    Call Foo(a, _', '            b)', 'End Sub'));
	});

	it('preserves CRLF line endings', () => {
		expect(format('Sub T()\r\nx = 1\r\nEnd Sub\r\n')).toBe('Sub T()\r\n    x = 1\r\nEnd Sub\r\n');
	});

	it('leaves Attribute lines exactly as written', () => {
		const src = lines('Attribute VB_Name = "Module1"', 'Sub T()', 'Attribute T.VB_Description = "x"', 'x=1', 'End Sub');
		expect(format(src)).toBe(lines('Attribute VB_Name = "Module1"', 'Sub T()', 'Attribute T.VB_Description = "x"', '    x = 1', 'End Sub'));
	});
});

describe('formatVbaModule - casing', () => {
	it('canonicalizes keyword casing', () => {
		expect(format(lines('option explicit', 'sub t()', 'dim x as long', 'if x then exit sub', 'end sub')))
			.toBe(lines('Option Explicit', 'Sub t()', '    Dim x As Long', '    If x Then Exit Sub', 'End Sub'));
	});

	it('cases a contextual keyword only inside its statement (issue #86)', () => {
		// The VBE writes `Option Compare Text` and keeps a variable called
		// text, binary or output as it was declared.
		const src = lines(
			'option compare text',
			'function describe(byval text as string, byval binary as long) as string',
			'dim output as string',
			'output = text & cstr(binary)',
			'open output for output as #1',
			'end function',
		);
		expect(format(src)).toBe(lines(
			'Option Compare Text',
			'Function describe(ByVal text As String, ByVal binary As Long) As String',
			'    Dim output As String',
			'    output = text & CStr(binary)',
			'    Open output For Output As #1',
			'End Function',
		));
	});

	it('leaves the lines a comment runs on to through _ exactly as written (issue #82)', () => {
		// The VBE reads `End Sub` and `this is not code` as comment text.
		const hidden = lines('Sub Test()', "' disabled: _", '  End Sub', 'Debug.Print 1', 'End Sub');
		expect(format(hidden)).toBe(lines('Sub Test()', "    ' disabled: _", '  End Sub', '    Debug.Print 1', 'End Sub'));
		const trailing = lines('Sub T()', "n = 1 ' note _", '    this is not code', 'n = 2', 'End Sub');
		expect(format(trailing)).toBe(lines('Sub T()', "    n = 1 ' note _", '    this is not code', '    n = 2', 'End Sub'));
	});

	it('leaves strings and comments alone', () => {
		const src = lines('Sub T()', 'x = "end sub if then"', "' if then else", 'Rem dim x', 'End Sub');
		expect(format(src)).toBe(lines('Sub T()', '    x = "end sub if then"', "    ' if then else", '    Rem dim x', 'End Sub'));
	});

	it('cases identifiers through the resolver, but not member names or named arguments', () => {
		const canonical = new Map([['myvar', 'myVar'], ['name', 'Name'], ['foo', 'Foo']]);
		const src = lines('Sub T()', 'MYVAR = obj.NAME', 'foo NAME:=1', 'End Sub');
		const out = format(src, { ...SPACES, identifierCase: (name) => canonical.get(name.toLowerCase()) });
		expect(out).toBe(lines('Sub T()', '    myVar = obj.NAME', '    Foo NAME:=1', 'End Sub'));
	});

	it('ignores a resolver that answers with a different word', () => {
		const src = lines('Sub T()', 'x = 1', 'End Sub');
		expect(format(src, { ...SPACES, identifierCase: () => 'y' })).toBe(lines('Sub T()', '    x = 1', 'End Sub'));
	});
});

describe('formatVbaModule - spacing', () => {
	function body(statement: string, options = SPACES): string {
		const out = format(lines('Sub T()', statement, 'End Sub'), options).split('\n');
		return out[1].trimStart();
	}

	it('spaces assignments and comparisons', () => {
		expect(body('x=1')).toBe('x = 1');
		expect(body('If a=b Then c=d')).toBe('If a = b Then c = d');
		expect(body('If a<>b And c<=d And e>=f And g<h Then')).toBe('If a <> b And c <= d And e >= f And g < h Then');
	});

	it('keeps a relational operator written the other way round whole, and spaces it (issue #87)', () => {
		expect(body('If a=>b And c=<d And e><f Then')).toBe('If a => b And c =< d And e >< f Then');
		expect(body('x = 1.')).toBe('x = 1.');
		expect(body('x=&17')).toBe('x = &17');
	});

	it('spaces after commas, semicolons and statement separators', () => {
		expect(body('Cells(1,2).Value=3')).toBe('Cells(1, 2).Value = 3');
		expect(body('Debug.Print "a";"b"')).toBe('Debug.Print "a"; "b"');
		expect(body('x = 1:y = 2')).toBe('x = 1: y = 2');
		expect(body('Get #1, ,rec')).toBe('Get #1, , rec');
	});

	it('spaces binary operators and leaves unary minus alone', () => {
		expect(body('x=a+b*c-d/e\\f^g')).toBe('x = a + b * c - d / e \\ f ^ g');
		expect(body('x=a+b&"c"')).toBe('x = a + b&"c"');
		expect(body('x=(a)&"c"')).toBe('x = (a) & "c"');
		expect(body('x=-1')).toBe('x = -1');
		expect(body('x=a*-b')).toBe('x = a * -b');
		expect(body('x=f(a)-1')).toBe('x = f(a) - 1');
		expect(body('x=(a+b)*c')).toBe('x = (a + b) * c');
		expect(body('x=-(a+b)')).toBe('x = -(a + b)');
		expect(body('x = y Mod-2')).toBe('x = y Mod -2');
		expect(body('For i=1 To n Step-1')).toBe('For i = 1 To n Step -1');
		expect(body('Debug.Print -1')).toBe('Debug.Print -1');
		expect(body('Foo-1')).toBe('Foo -1');
		expect(body('obj.Method -1, 2')).toBe('obj.Method -1, 2');
		expect(body('Debug.Print a-1')).toBe('Debug.Print a - 1');
		expect(body('Case -1')).toBe('Case -1');
		expect(body('x = [A1]+1')).toBe('x = [A1] + 1');
		expect(body('x = "a"&"b"')).toBe('x = "a" & "b"');
		expect(body('x = Date-1')).toBe('x = Date - 1');
		expect(body('x = o.Type+1')).toBe('x = o.Type + 1');
	});

	it('never removes existing spaces, so aligned code survives', () => {
		expect(body('x    =    1')).toBe('x    =    1');
		expect(body("x = 1      ' aligned")).toBe("x = 1      ' aligned");
		expect(body('Const A   As Long = 1')).toBe('Const A   As Long = 1');
	});

	it('keeps type-declaration characters and bangs glued to their names', () => {
		expect(body('Dim s$, n&, d#, f!, c@')).toBe('Dim s$, n&, d#, f!, c@');
		expect(body('s$=Left$(t$,1)')).toBe('s$ = Left$(t$, 1)');
		expect(body('x = rs!Field')).toBe('x = rs!Field');
		expect(body('x = a&b')).toBe('x = a&b');
	});

	it('does not space named arguments or the file-number hash', () => {
		expect(body('Foo a:=1,b:=2')).toBe('Foo a:=1, b:=2');
		expect(body('Open f For Input As #1')).toBe('Open f For Input As #1');
		expect(body('Print #1,x')).toBe('Print #1, x');
	});

	it('drops trailing whitespace but keeps a continuation underscore', () => {
		const src = lines('Sub T()', 'x = 1   ', 'y = a + _', 'b', 'End Sub');
		expect(format(src)).toBe(lines('Sub T()', '    x = 1', '    y = a + _', '        b', 'End Sub'));
	});

	it('continues a line whose underscore has spaces after it, and drops the spaces (issue #83)', () => {
		// The VBE reads `_   ` as a line continuation and drops the spaces
		// when it stores the line.
		const src = lines('Sub T()', 'y = a + _   ', 'b', 'End Sub');
		expect(format(src)).toBe(lines('Sub T()', '    y = a + _', '        b', 'End Sub'));
	});

	it('keeps an underscore with no space before it off the next line', () => {
		// `&_` is no continuation (MS-VBAL 3.2.2 wants whitespace before the
		// underscore), so the next line stays a statement of its own.
		const src = lines('Sub T()', 'y = a &_  ', 'b', 'End Sub');
		expect(format(src)).toBe(lines('Sub T()', '    y = a &_', '    b', 'End Sub'));
	});

	it('spaces a #Const value', () => {
		expect(format('#Const A=1')).toBe('#Const A = 1');
	});
});

describe('formatVbaModule - safety', () => {
	it('is idempotent', () => {
		const src = lines('sub t()', 'if a=b then', 'x=-1', 'end if', 'end sub');
		const once = format(src);
		expect(format(once)).toBe(once);
	});

	it('reports the first token difference between two sources', () => {
		expect(tokenStreamDifference('x = 1', 'x = 2')).toContain('changed from "1" to "2"');
		expect(tokenStreamDifference('x = 1', 'X = 1')).toBeUndefined();
		expect(tokenStreamDifference('x = 1', 'x = 1 + 2')).toContain('token count');
		expect(tokenStreamDifference('x = 1 + _\n2', 'x = 1 + 2')).toContain('continuation');
	});

	it('refuses to move a line continuation, even when the count stays the same', () => {
		// Placed at column 1, the leading ` _` lost its space and became a stray
		// `_`, and the stray `_` below it gained a tab and became a continuation:
		// one continuation either way, a different module (found by
		// tests/properties).
		expect(tokenStreamDifference(' _\r\n_\r\n', '_\r\n\t_\r\n')).toContain('moved');
		const result = formatVbaModule(' _\r\n_\r\n', { tabSize: 0, insertSpaces: false });
		expect(result.text).toBeUndefined();
	});

	it('formats an empty module and a module of only blank lines', () => {
		expect(format('')).toBe('');
		expect(format('\n\n')).toBe('\n\n');
	});
});
