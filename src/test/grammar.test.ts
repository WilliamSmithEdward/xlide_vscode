// The VBA grammar in a real VS Code: the TextMate scopes VS Code computes for
// a loose module, read through the command its own colorizer tests use.

import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { activate, workspaceRoot } from './support';

interface SyntaxToken {
	/** The token's text. */
	c: string;
	/** Its scopes, space-separated. */
	t: string;
}

/** The tokens of `file`, as VS Code colors it. */
async function tokensOf(file: string): Promise<SyntaxToken[]> {
	const tokens = await vscode.commands.executeCommand<SyntaxToken[]>(
		'_workbench.captureSyntaxTokens', vscode.Uri.file(file),
	);
	return tokens ?? [];
}

/** The scopes of the first token holding `text` in `file`. */
async function scopesOf(file: string, text: string): Promise<string> {
	const token = (await tokensOf(file)).find((candidate) => candidate.c.includes(text));
	assert.ok(token, `no token holds ${JSON.stringify(text)}`);
	return token.t;
}

/** The scopes of the first token that is exactly `text` in `file`. */
async function scopesOfToken(file: string, text: string): Promise<string> {
	const token = (await tokensOf(file)).find((candidate) => candidate.c === text);
	assert.ok(token, `no token is ${JSON.stringify(text)}`);
	return token.t;
}

/** Writes `lines` to a loose module, runs `check` on it and removes it. */
async function withProbe(lines: string[], check: (file: string) => Promise<void>): Promise<void> {
	const file = path.join(workspaceRoot(), 'GrammarProbe.bas');
	fs.writeFileSync(file, [...lines, ''].join('\r\n'));
	try {
		await check(file);
	} finally {
		fs.rmSync(file, { force: true });
	}
}

/**
 * Writes `lines` to a loose module and returns, for each line, the tokens that
 * start with a letter, as `text innermost-scope`, so a test reads exactly the
 * tokens of the line it names.
 */
async function lineTokens(lines: string[]): Promise<string[][]> {
	const file = path.join(workspaceRoot(), 'GrammarProbe.bas');
	fs.writeFileSync(file, [...lines, ''].join('\r\n'));
	try {
		const tokens = await tokensOf(file);
		let next = 0;
		return lines.map((line) => {
			const words: string[] = [];
			let text = '';
			while (text.length < line.length && next < tokens.length) {
				const token = tokens[next++];
				text += token.c;
				if (/^\p{L}/u.test(token.c.trim())) {
					words.push(`${token.c.trim()} ${token.t.split(' ').at(-1)}`);
				}
			}
			assert.equal(text, line, 'the tokens follow the lines');
			return words;
		});
	} finally {
		fs.rmSync(file, { force: true });
	}
}

const KEYWORD = 'keyword.control.vba';
const NAME = 'variable.other.vba';
const MEMBER = 'variable.other.member.vba';

suite('VBA grammar', () => {
	suiteSetup(async () => {
		await activate();
	});

	test('colors the line a comment ending in _ runs on to as comment text (issue #82)', async () => {
		const file = path.join(workspaceRoot(), 'GrammarProbe.bas');
		fs.writeFileSync(file, [
			'Sub Test()',
			"    ' disabled: _",
			'    End Sub',
			'    Debug.Print 1',
			"#If True Then ' note _",
			'    not code',
			'#End If',
			'    x = a + _   ',
			'        b',
			'End Sub',
			'',
		].join('\r\n'));
		try {
			assert.match(await scopesOf(file, 'End Sub'), /\bcomment\.line\.apostrophe\.vba\b/);
			assert.match(await scopesOf(file, 'not code'), /\bcomment\.line\.apostrophe\.vba\b/);
			assert.doesNotMatch(await scopesOf(file, 'Debug'), /comment/);
			assert.match(await scopesOf(file, ' _   '), /\bpunctuation\.separator\.continuation\.vba\b/);
		} finally {
			fs.rmSync(file, { force: true });
		}
	});

	// Excel accepts these: oracle issue682_16_runtime (`Print #f, "abc": Close #f`),
	// issue146_a_runtime (`Close #&H1, #&H2`), issue239_17_runtime (`#12/30/9999# + 1`),
	// issue1416_is_cvdate_1_1_2026_compile, and issue190_16_compile, issue190_20_compile,
	// issue1323_11_runtime and issue133_q_runtime (the first `d =` line); Word, PowerPoint
	// and Access accept issue1416_<host>_is_cvdate_1_1_2026_compile. The oracle has no case
	// for the dates of the second `d =` line: `#8:30#`, `#1 PM#` and parts separated by
	// U+3000 (whitespace to VBA, MS-VBAL 3.2.2) are date-or-time forms of MS-VBAL 3.3.3,
	// which the lexer reads too.
	test('colors # signs as a date only when the text between them is a date', async () => {
		await withProbe([
			'Sub Test()',
			'    f = FreeFile: Open p For Output As #f: Print #f, "abc": Close #f',
			'    Close #&H1, #&H2',
			'    t = (86400# - s) * 1000#',
			'    Main = #12/30/9999# + 1',
			'    Debug.Print CVDate(#1/1/2026#) Is Nothing',
			'    d = #Jan 1, 2000# + #2000-02-28# + #1/1/2000 12:00:00 AM# + #13:00 PM#',
			'    d = #8:30# + #1 PM# + #1\u30001\u30002024#',
			'End Sub',
		], async (file) => {
			const dates = (await tokensOf(file))
				.filter((token) => /\bconstant\.other\.date\.vba\b/.test(token.t))
				.map((token) => token.c);
			assert.deepEqual(dates, [
				'#12/30/9999#', '#1/1/2026#', '#Jan 1, 2000#', '#2000-02-28#', '#1/1/2000 12:00:00 AM#',
				'#13:00 PM#', '#8:30#', '#1 PM#', '#1\u30001\u30002024#',
			]);
			assert.match(await scopesOf(file, 'abc'), /\bstring\.quoted\.double\.vba\b/);
			assert.match(await scopesOfToken(file, '&H1'), /\bconstant\.numeric\.hex\.vba\b/);
			assert.match(await scopesOfToken(file, '86400#'), /\bconstant\.numeric\.float\.vba\b/);
		});
	});

	// Excel accepts these: oracle issue681_09_runtime, issue681_11_runtime and
	// issue234_01_compile. The VBE ends a string at the end of its line.
	test('ends a string left open at the end of its line', async () => {
		await withProbe([
			'Private Const K As String = "abc',
			'Function Main() As Variant',
			'    Main = "abc_',
			'    Main = Main & "|"',
			'End Function',
			'#If False Then',
			'asdf "qwer',
			'#End If',
		], async (file) => {
			assert.match(await scopesOfToken(file, 'Function'), /\bstorage\.type\.function\.vba\b/);
			assert.match(await scopesOf(file, 'abc_'), /\bstring\.quoted\.double\.vba\b/);
			assert.match(await scopesOf(file, '#End If'), /\bkeyword\.control\.preprocessor\.vba\b/);
		});
	});

	// Excel accepts these: oracle dim_type_suffix_only_compile, suffix_long_amp_inrange_value_compile,
	// issue133_q_runtime, issue232_27_runtime, issue1412_l_dim_long_max_compile,
	// issue1412_l_redim_double_in_range_runtime and issue1412_d_divide_ok_compile; Word,
	// PowerPoint and Access accept the same three issue1412_<host>_… cases. The oracle has no
	// case for `1D3`, `.5` or `1.`: they are the exponent letter D and the optional digits
	// around the point of MS-VBAL 3.3.2, which the lexer reads too.
	test('colors a number with its type suffix or exponent as one number', async () => {
		await withProbe([
			'Sub Test()',
			'    Dim a(0 To 2147483647&) As Byte',
			'    ReDim b(10#)',
			'    Dim c(4294967294# / 2) As Long',
			'    price = 12.34@: ratio = 0.5#: flag = 1%: total = 6@',
			'    x = 40000&',
			'    v = 1E+308: v = 3.402823E+38!: v = CLngLng(9E+18): v = 9223372036854775807^',
			'    v = 1e-3: v = 1D3: v = .5: v = 1.',
			'End Sub',
		], async (file) => {
			for (const integer of ['2147483647&', '40000&', '1%', '9223372036854775807^']) {
				assert.match(await scopesOfToken(file, integer), /\bconstant\.numeric\.integer\.vba\b/, integer);
			}
			for (const float of [
				'10#', '4294967294#', '12.34@', '0.5#', '6@', '1E+308', '3.402823E+38!', '9E+18', '1e-3', '1D3', '.5', '1.',
			]) {
				assert.match(await scopesOfToken(file, float), /\bconstant\.numeric\.float\.vba\b/, float);
			}
		});
	});

	// Excel accepts these: oracle const_directive_before_option_explicit_compile,
	// issue192_12_runtime, issue192_22_runtime and issue208_12_runtime.
	test('colors the numbers, dates and strings of a directive line', async () => {
		await withProbe([
			'#Const XLIDEORACLEFLAG = 1',
			'#Const S = "x"',
			'#If 5 / 2 = 2.5 Then',
			'#ElseIf #1/2/2000# > #1/1/2000# Then',
			'#End If',
		], async (file) => {
			assert.match(await scopesOfToken(file, '1'), /\bmeta\.preprocessor\.vba constant\.numeric\.integer\.vba\b/);
			assert.match(await scopesOfToken(file, '2.5'), /\bmeta\.preprocessor\.vba constant\.numeric\.float\.vba\b/);
			assert.match(await scopesOfToken(file, '#1/2/2000#'), /\bmeta\.preprocessor\.vba constant\.other\.date\.vba\b/);
			assert.match(await scopesOf(file, 'x'), /\bmeta\.preprocessor\.vba string\.quoted\.double\.vba\b/);
		});
	});

	test('colors a word that is a keyword only in its statement by where it stands (issue #1575)', async () => {
		const OTHER = 'keyword.other.vba';
		const FUNCTION = 'support.function.builtin.vba';
		const lines = await lineTokens([
			'Open p For Binary Access Read Write Lock Read As #2',
			'Open p For Random Shared As #3',
			'Option Compare Text',
			'Declare PtrSafe Function F CDecl Lib "k" Alias "G" () As Long',
			'On Error GoTo Error',
			'Error:',
			'If x Then Error 5: Error 6',
			'10 Error 7',
			'Main = Error(-1) + InputB(1) + Circle(1)',
			'Error = Error + 1',
			'Dim Step As Long, Lib As Long, Alias As Long',
			'For i = 1 To Step Step 2: Step = 1: Lib = Alias: Next',
			'For Output = 1 To 2',
			'    Step 3',
			'Output = Binary + Random + Read + Text',
		]);
		const words = (i: number, scope: string): string[] => lines[i].filter((word) => word.endsWith(` ${scope}`)).map((word) => word.slice(0, -scope.length - 1));
		assert.deepEqual(words(0, KEYWORD), ['Open', 'For', 'Binary', 'Access', 'Read Write', 'Lock', 'Read']);
		assert.deepEqual(words(1, KEYWORD), ['Open', 'For', 'Random']);
		assert.deepEqual(lines[2], ['Option keyword.control.option.vba', 'Compare Text keyword.other.option.vba']);
		assert.deepEqual(words(3, OTHER), ['CDecl', 'Lib', 'Alias', 'As']);
		assert.deepEqual(lines[4], [`On ${KEYWORD}`, `Error ${KEYWORD}`, `GoTo ${KEYWORD}`, `Error ${NAME}`]);
		assert.deepEqual(lines[5], [`Error ${NAME}`]);
		assert.deepEqual(words(6, KEYWORD), ['If', 'Then', 'Error', 'Error']);
		assert.deepEqual(words(7, KEYWORD), ['Error']);
		assert.deepEqual(words(8, FUNCTION), ['Error', 'InputB', 'Circle']);
		assert.deepEqual(lines[9], [`Error ${NAME}`, `Error ${NAME}`]);
		assert.deepEqual(words(10, NAME), ['Step', 'Lib', 'Alias']);
		assert.deepEqual(words(11, KEYWORD), ['For', 'To', 'Step', 'Next']);
		assert.deepEqual(words(11, NAME), ['i', 'Step', 'Step', 'Lib', 'Alias']);
		assert.deepEqual(words(12, NAME), ['Output']);
		assert.deepEqual(words(13, KEYWORD), ['Step']);
		assert.deepEqual(words(14, NAME), ['Output', 'Binary', 'Random', 'Read', 'Text']);
	});

	// Excel accepts these: oracle legacy_gosub_return_valid_probe, issue240_19_runtime and
	// issue1411_e_exit_between_runtime (Exit Sub), exit_function_in_property_get (Exit
	// Function); Word, PowerPoint and Access accept issue1411_<host>_e_exit_between_runtime
	// (Exit Sub). The oracle has no case for Exit Property, the Exit Property statement of
	// MS-VBAL 5.4.2.19:
	// https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/2c254f33-48b5-40b7-9dc7-4943313a1742
	test('colors the procedure word of Exit Sub, Exit Function and Exit Property as a keyword', async () => {
		const lines = await lineTokens([
			'Sub Test()',
			'    If n > 3 Then Exit Sub',
			'    Exit Sub',
			'End Sub',
			'Function F() As Long',
			'    Exit Function',
			'End Function',
			'Property Get P() As Long',
			'    Exit Property',
			'End Property',
			'    Exit For: Exit Do',
			'    ExitSub = Exit_Sub',
		]);
		assert.deepEqual(lines[1], [`If ${KEYWORD}`, `n ${NAME}`, `Then ${KEYWORD}`, `Exit ${KEYWORD}`, `Sub ${KEYWORD}`]);
		assert.deepEqual(lines[2], [`Exit ${KEYWORD}`, `Sub ${KEYWORD}`]);
		assert.deepEqual(lines[5], [`Exit ${KEYWORD}`, `Function ${KEYWORD}`]);
		assert.deepEqual(lines[8], [`Exit ${KEYWORD}`, `Property ${KEYWORD}`]);
		assert.deepEqual(lines[10], [`Exit ${KEYWORD}`, `For ${KEYWORD}`, `Exit ${KEYWORD}`, `Do ${KEYWORD}`]);
		assert.deepEqual(lines[11], [`ExitSub ${NAME}`, `Exit_Sub ${NAME}`]);
	});

	test('colors a name after . or ! as a member, whatever it is spelled like (issue #1574)', async () => {
		const lines = await lineTokens([
			'ws.Range("A1").End(xlUp).Select',
			'wb.Close: doc.Close: pres.Close: db.Close',
			'Debug.Print rs!Date',
			'With rs: .Close: x = !Date: End With',
			's = re.Replace(t, shp.Left)',
			'n = VBA.Len(s) + Excel.xlUp',
			'x! = y! + Len(s)',
			'If True Then Exit Sub Else End',
		]);
		const members = lines.map((line) => line.filter((word) => word.endsWith(MEMBER)).map((word) => word.split(' ')[0]));
		assert.deepEqual(members, [
			['Range', 'End', 'Select'],
			['Close', 'Close', 'Close', 'Close'],
			['Print', 'Date'],
			['Close', 'Date'],
			['Replace', 'Left'],
			['Len', 'xlUp'],
			[],
			[],
		]);
		assert.deepEqual(lines[3].slice(0, 2), [`With ${KEYWORD}`, `rs ${NAME}`]);
		assert.deepEqual(lines[6], [`x ${NAME}`, `y ${NAME}`, 'Len support.function.builtin.vba', `s ${NAME}`]);
		assert.deepEqual(lines[7], [`If ${KEYWORD}`, 'True keyword.other.literal.vba', `Then ${KEYWORD}`, `Exit ${KEYWORD}`, `Sub ${KEYWORD}`, `Else ${KEYWORD}`, `End ${KEYWORD}`]);
	});

	// Excel accepts these: oracle issue262_70_runtime (`Reset` then `Line Input #7, s`),
	// issue682_23_runtime (`Name p As p & ".new"`) and issue146_a_runtime (`Reset`). The
	// oracle has no accepted case with Width #, and in Word, PowerPoint and Access only
	// rejected ones (issue1415_<host>_g5_line_input_number_array_compile, a type mismatch),
	// where these are the same statements: MS-VBAL 5.4.5.6 (Line Input),
	// https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/dab5496c-a151-4d69-adc4-bb5effc066e9
	// 5.4.5.7 (Width),
	// https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/e427933b-d398-424b-9dbc-8cb91dece4cc
	// 5.4.5.2 (Reset, a close-statement),
	// https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/73ef1ac2-4da7-4cda-b2a3-5984a8649ded
	// and the Name statement of the Office VBA reference.
	// https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/name-statement
	// The VBE colors Line and Input of Line Input, and Name and As of Name ... As, as
	// keywords, but leaves Width (in `Width #1, 80`) and Reset as names, in Excel, Word,
	// PowerPoint and Access alike, so the grammar does too.
	test('colors Line Input and Name ... As as statements, and Width # and Reset as names', async () => {
		const lines = await lineTokens([
			'Sub Test()',
			'    Reset',
			'    Line Input #7, s: Width #1, 80: Width #f, 80: Width #(f), 80',
			'    Name p As p & ".new"',
			'    Name "C:\\old.txt" As "C:\\new.txt"',
			'    If EOF(7) Then Reset: Name a As b Else Reset',
			'10  Reset',
			'    Name = ws.Name: w = ws.Width: GoTo Width: GoTo Reset',
			'    Dim Name As String, Reset As Long: Reset = 1: Reset% = 2: Reset^ = 3: Width# = 4',
			'    Line (0, 0)-(1, 1): Width = 3: Consume1 Reset: Name a \' As b',
			'Reset:',
			'End Sub',
			'Private Type Rec',
			'    Name As String * 3',
			'    Reset As Long',
			'End Type',
		]);
		const keywords = lines.map((line) => line.filter((word) => word.endsWith(KEYWORD)).map((word) => word.split(' ')[0]));
		assert.deepEqual(keywords.slice(1, 7), [
			[],
			['Line', 'Input'],
			['Name'],
			['Name'],
			['If', 'Then', 'Name', 'Else'],
			[],
		]);
		const widthAndReset = lines.slice(1, 7).map((line) => line.filter((word) => /^(Width|Reset) /.test(word)));
		assert.deepEqual(widthAndReset, [
			[`Reset ${NAME}`],
			[`Width ${NAME}`, `Width ${NAME}`, `Width ${NAME}`],
			[],
			[],
			[`Reset ${NAME}`, `Reset ${NAME}`],
			[`Reset ${NAME}`],
		]);
		assert.deepEqual(lines[7], [
			`Name ${NAME}`, `ws ${NAME}`, `Name ${MEMBER}`, `w ${NAME}`, `ws ${NAME}`, `Width ${MEMBER}`,
			`GoTo ${KEYWORD}`, `Width ${NAME}`, `GoTo ${KEYWORD}`, `Reset ${NAME}`,
		]);
		assert.deepEqual(keywords.slice(8, 11), [[], [], []]);
		assert.deepEqual(lines[13], [`Name ${NAME}`, 'As keyword.other.vba', 'String support.type.vba']);
		assert.deepEqual(lines[14], [`Reset ${NAME}`, 'As keyword.other.vba', 'Long support.type.vba']);
	});

	// Issue #1617. The VBE shows Me, the vb and xl constants, Date and Object as values and
	// String, Seek and Input as calls as names, in Excel, Word, PowerPoint and Access. Date,
	// String and Object stay types after As, and Seek and Input stay keywords in their
	// statements: Seek with or without # (MS-VBAL 5.4.5.3), Input # (5.4.5.10).
	test('colors Me, the constants and the calls the VBE shows as names as names', async () => {
		const TYPE = 'support.type.vba';
		const lines = await lineTokens([
			'    Set ws = Me.Worksheets(1)',
			'    s = "a" & vbCrLf & vbNullString: d = xlUp',
			'    x = Date: Set o = Object: y = Date$',
			'    s = String(2, "a") & String$(2, "b")',
			'    p = Seek(1): t = Input(4, #1) & Input$(4, #1)',
			'    Dim a As String, b As Date, c As Object, n As Long',
			'    Seek #1, 1: Seek 1, 1',
			'    Input #1, x',
			'    If p Then Seek 1, 2 Else Seek #1, 3',
		]);
		assert.deepEqual(lines[0], [`Set keyword.other.vba`, `ws ${NAME}`, `Me ${NAME}`, `Worksheets ${MEMBER}`]);
		assert.deepEqual(lines[1].slice(1), [`vbCrLf ${NAME}`, `vbNullString ${NAME}`, `d ${NAME}`, `xlUp ${NAME}`]);
		assert.deepEqual(lines[2].filter((word) => /^(Date|Object)\b/.test(word)), [`Date ${NAME}`, `Object ${NAME}`, `Date ${NAME}`]);
		assert.deepEqual(lines[3].filter((word) => word.startsWith('String')), [`String ${NAME}`, `String ${NAME}`]);
		assert.deepEqual(lines[4].filter((word) => /^(Seek|Input)\b/.test(word)), [`Seek ${NAME}`, `Input ${NAME}`, `Input ${NAME}`]);
		assert.deepEqual(lines[5].filter((word) => word.endsWith(TYPE)), [`String ${TYPE}`, `Date ${TYPE}`, `Object ${TYPE}`, `Long ${TYPE}`]);
		assert.deepEqual(lines[6].filter((word) => word.startsWith('Seek')), [`Seek ${KEYWORD}`, `Seek ${KEYWORD}`]);
		assert.deepEqual(lines[7][0], `Input ${KEYWORD}`);
		assert.deepEqual(lines[8].filter((word) => word.startsWith('Seek')), [`Seek ${KEYWORD}`, `Seek ${KEYWORD}`]);
	});

	// Oracle issue231_01/07/09/12/13_runtime and issue231_60/61_compile
	// cover one-line If, colon, label, line number and continued Rem comments.
	test('colors statement-following Rem and its continuation as comments', async () => {
		await withProbe([
			'If True Then Main = 1 Rem first', 'If False Then Main = 1 Rem rest Else Main = 2',
			'If True Then Main = 1 Else Rem branch', 'Main = 1: Rem colon',
			'L1: Rem label', '10 Rem numbered', 'Main = 1: Rem note _', 'continued',
			'wb.Save: Rem Excel', 'doc.Save: Rem Word',
			'pres.Save: Rem PowerPoint', 'rs.MoveNext: Rem Access',
		], async (file) => {
			for (const text of ['Rem first', 'Rem rest Else Main = 2', 'Rem branch', 'Rem colon',
				'Rem label', 'Rem numbered', 'continued', 'Rem Excel', 'Rem Word', 'Rem PowerPoint', 'Rem Access']) {
				assert.match(await scopesOfToken(file, text), /comment\.line\.rem\.vba/);
			}
		});
	});

	// Oracle issue231_59_compile rejects Rem immediately after Then.
	test('keeps Then Rem as words and Rem members as members', async () => {
		await withProbe(['If True Then Rem note'], async (file) => {
			assert.match(await scopesOfToken(file, 'Rem'), /keyword\.other\.vba/);
			assert.doesNotMatch(await scopesOfToken(file, 'note'), /comment/);
		});
		await withProbe(['x.Rem = 1', 'x!Rem = 2'], async (file) => {
			assert.match(await scopesOfToken(file, 'Rem'), /variable\.other\.member\.vba/);
		});
	});

	// Oracle issue1323_06_runtime and issue1321f_14/15_runtime cover colons;
	// issue1331_02/03/10_runtime only compile inside a skipped arm. They do
	// not authorize recognizing directives after a colon.
	test('ends a directive scope at its own colon, outside literals and brackets', async () => {
		for (const line of ['#Const A = 1', '#Else', '#ElseIf True Then', '#End If']) {
			await withProbe([line + ': Rem tail'], async (file) => {
				assert.match(await scopesOfToken(file, 'Rem tail'), /comment\.line\.rem\.vba/);
				assert.doesNotMatch(await scopesOf(file, ':'), /meta\.preprocessor/);
			});
		}
		await withProbe(['#Const A = 1: total = 2', 'x = 1: #Const B = 2'], async (file) => {
			assert.match(await scopesOfToken(file, 'total'), /variable\.other\.vba/);
			assert.doesNotMatch(await scopesOfToken(file, 'Const'), /preprocessor/);
		});
		// Existing date/string coloring must survive the directive boundary.
		await withProbe(['#If #12:00:00 PM# = 0.5 Then', '#Const A = "a:b"',
			'#Const [Rem] = 1', '#Const [a:b] = 2'], async (file) => {
			assert.match(await scopesOfToken(file, '#12:00:00 PM#'), /constant\.other\.date\.vba/);
			assert.match(await scopesOf(file, 'Then'), /meta\.preprocessor/);
			assert.match(await scopesOf(file, 'a:b'), /string\.quoted/);
			assert.match(await scopesOf(file, '[Rem]'), /meta\.preprocessor/);
			assert.match(await scopesOf(file, '[a:b]'), /meta\.preprocessor/);
		});
	});

	// Every host compiles accented Latin names (oracle nonascii_names_<host>_runtime). The other
	// scripts follow the lexer, since the test projects' code page holds Latin only.
	test('colors a name outside ASCII as one name in every rule that reads a name (issue #1604)', async () => {
		const lines = await lineTokens([
			'Attribute État.VB_VarUserMemId = 0',
			'Implements IÉtat',
			'Public Type Größe',
			'Private Enum Çouleur',
			'Public Event Geändert(ByVal 値 As Long)',
			'Private Declare PtrSafe Sub Zähler Lib "x" ()',
			'Private Property Get Länge() As Long',
			'Function 名前() As Long',
			'Sub Ячейка()',
			'    Dim ค่า As Long, IfÉtat As Long, cafe\u0301 As Long',
			'    ThisWorkbook.Größe = ActiveDocument.Länge + ActivePresentation.Значение + CurrentDb!ค่า',
			'    If値 = 値If',
			'End Sub',
		]);
		const name = (word: string) => word.split(' ')[0];
		const scope = (word: string) => word.split(' ')[1];
		assert.deepEqual(lines[0], ['Attribute keyword.other.attribute.vba', 'État.VB_VarUserMemId variable.other.attribute.vba']);
		assert.deepEqual(lines[1], ['Implements keyword.other.implements.vba', 'IÉtat entity.name.type.vba']);
		assert.deepEqual(lines[2], ['Public storage.modifier.vba', 'Type storage.type.vba', 'Größe entity.name.type.vba']);
		assert.deepEqual(lines[3], ['Private storage.modifier.vba', 'Enum storage.type.vba', 'Çouleur entity.name.type.vba']);
		assert.deepEqual(lines[4].slice(0, 3), ['Public storage.modifier.vba', 'Event keyword.other.event.vba', 'Geändert entity.name.function.vba']);
		assert.ok(lines[4].includes(`値 ${NAME}`), lines[4].join());
		assert.ok(lines[5].includes('Zähler entity.name.function.vba'), lines[5].join());
		assert.ok(lines[6].includes('Länge entity.name.function.property.vba'), lines[6].join());
		assert.deepEqual(lines[7].slice(0, 2), ['Function storage.type.function.vba', '名前 entity.name.function.vba']);
		assert.deepEqual(lines[8], ['Sub storage.type.function.vba', 'Ячейка entity.name.function.vba']);
		assert.deepEqual(lines[9].filter((word) => scope(word) === NAME).map(name), ['ค่า', 'IfÉtat', 'cafe\u0301']);
		assert.deepEqual(lines[10].filter((word) => scope(word) === MEMBER).map(name), ['Größe', 'Länge', 'Значение', 'ค่า']);
		assert.deepEqual(lines[11], [`If値 ${NAME}`, `値If ${NAME}`]);
	});

	// The VBE shows True, False, Nothing, Null and Empty as keywords in Excel, Word,
	// PowerPoint and Access, wherever a value can stand, and as names when they name a
	// Type field. They are the literal identifiers of MS-VBAL 3.3.5.2.
	// https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/7df907cb-ab6c-40d3-aa81-272742ce00c3
	test('colors True, False, Nothing, Null and Empty as keywords, and other constants as before', async () => {
		const LITERAL = 'keyword.other.literal.vba';
		const lines = await lineTokens([
			'Sub Test(Optional ByVal b As Boolean = True, Optional v As Variant = Empty)',
			'    x = False: Set o = Nothing: v = Null',
			'    If o Is Nothing Or IsNull(Null) Then x = TRUE',
			'    Const K As Boolean = false',
			'    s = vbCrLf & vbNullString: d = xlUp',
			'    x = a.True + a.Nothing + Truex + xNull',
			'    x = [True] + [Nothing]',
			'End Sub',
			'Private Type Flags',
			'    True As Long',
			'    Nothing(1 To 2) As Long',
			'End Type',
		]);
		const literals = lines.map((line) => line.filter((word) => word.endsWith(` ${LITERAL}`)).map((word) => word.split(' ')[0]));
		assert.deepEqual(literals.slice(0, 4), [['True', 'Empty'], ['False', 'Nothing', 'Null'], ['Nothing', 'Null', 'TRUE'], ['false']]);
		// The constants are names since issue #1617.
		assert.deepEqual(lines[4].filter((word) => /^(vbCrLf|vbNullString|xlUp) /.test(word)), [
			`vbCrLf ${NAME}`, `vbNullString ${NAME}`, `xlUp ${NAME}`,
		]);
		assert.deepEqual(lines[5], [`x ${NAME}`, `a ${NAME}`, `True ${MEMBER}`, `a ${NAME}`, `Nothing ${MEMBER}`, `Truex ${NAME}`, `xNull ${NAME}`]);
		assert.deepEqual(lines[6], [`x ${NAME}`]);
		assert.deepEqual(literals.slice(9, 11), [[], []]);
		assert.deepEqual(lines[9][0], `True ${NAME}`);
		assert.deepEqual(lines[10][0], `Nothing ${NAME}`);
		await withProbe(['    x = [True] + [Nothing]'], async (file) => {
			for (const bracketed of ['[True]', '[Nothing]']) {
				assert.equal((await scopesOfToken(file, bracketed)).split(' ').at(-1), 'variable.other.foreign.vba', bracketed);
			}
		});
	});

	// The VBE shows True, False, Nothing, Null and Empty as keywords inside #If, #ElseIf
	// and #Const lines too, in Excel, Word, PowerPoint and Access. The grammar gives them
	// a keyword scope there and leaves the rest of the expression in the directive scope.
	// They are the literal identifiers of MS-VBAL 3.3.5.2.
	// https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/7df907cb-ab6c-40d3-aa81-272742ce00c3
	test('colors True, False, Nothing, Null and Empty as keywords in a directive', async () => {
		const LITERAL = 'keyword.other.literal.vba';
		const lines = await lineTokens([
			'#Const A = True',
			'#If A = TRUE And B = false Then \' True',
			'#ElseIf Not Nothing Or Null = Empty Then',
			'#Else',
			'#End If',
			'#Const [True] = 1',
			'#Const Truex = 1',
		]);
		const literals = lines.map((line) => line.filter((word) => word.endsWith(` ${LITERAL}`)).map((word) => word.split(' ')[0]));
		assert.deepEqual(literals, [['True'], ['TRUE', 'false'], ['Nothing', 'Null', 'Empty'], [], [], [], []]);
		const DIRECTIVE = 'meta.preprocessor.vba';
		assert.deepEqual(lines[0], [`A = ${DIRECTIVE}`, `True ${LITERAL}`]);
		assert.deepEqual(lines[1], [`A = ${DIRECTIVE}`, `TRUE ${LITERAL}`, `And B = ${DIRECTIVE}`, `false ${LITERAL}`, `Then ${DIRECTIVE}`]);
		assert.deepEqual(lines[5], []);
		assert.deepEqual(lines[6], [`Truex = ${DIRECTIVE}`]);
		await withProbe(['#If A Then \' True', '#Const [True] = 1'], async (file) => {
			const tokens = await tokensOf(file);
			const comment = tokens.find((token) => token.c.includes('True') && !token.c.includes('['));
			assert.ok(comment, 'the comment holds True');
			assert.match(comment.t.split(' ').at(-1) ?? '', /^comment\./);
			assert.equal((await scopesOf(file, '[True]')).split(' ').at(-1), DIRECTIVE);
		});
	});

});
