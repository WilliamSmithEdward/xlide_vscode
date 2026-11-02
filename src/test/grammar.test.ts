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
				if (/^[A-Za-z]/.test(token.c.trim())) {
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
		assert.deepEqual(lines[7], [`If ${KEYWORD}`, 'True constant.language.vba', `Then ${KEYWORD}`, `Exit ${KEYWORD}`, `Sub ${KEYWORD}`, `Else ${KEYWORD}`, `End ${KEYWORD}`]);
	});
});
