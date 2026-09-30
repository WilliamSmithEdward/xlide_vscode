// Arguments computed from InStr, InStrRev, Len, Asc and AscW of known strings
// (issue #201). The measured cases ran in Excel 16.0 through pyVBAharness:
// the raising ones raise error 5 every time they run, the others run.

import { describe, it, expect } from 'vitest';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';

function argumentErrors(body: string, option = '', extra = ''): string[] {
	const source = `Option Explicit\n${option ? `${option}\n` : ''}Function Main() As String\n    Dim s As String, p As String\n    ${body}\nEnd Function\n${extra}`;
	return analyzeVbaModuleSource({ source, moduleName: 'Module1' }).diagnostics
		.filter((d) => d.code === 'runtime-argument-value')
		.map((d) => d.message);
}

describe('arguments from InStr, Len and Asc of known strings (issue #201)', () => {
	const RAISES: ReadonlyArray<readonly [string, string, string]> = [
		['Left$ past a missing space', 's = "nospace"\n    Main = Left$(s, InStr(s, " ") - 1)', "Argument 'Length' of 'Left$' is -1"],
		['InStrRev with no separator', 'p = "file.txt"\n    Main = Left$(p, InStrRev(p, "\\") - 1)', "Argument 'Length' of 'Left$' is -1"],
		['Mid$ from a missing letter', 's = "abc"\n    Main = Mid$(s, InStr(s, "z"))', "Argument 'Start' of 'Mid$' is 0"],
		['Space$ of a longer string', 's = "abcdef"\n    Main = s & Space$(4 - Len(s))', "is -2"],
		['String$ of a longer string', 's = "abcdef"\n    Main = String$(5 - Len(s), "0") & s', "is -1"],
		['all literals', 'Main = Left$("nospace", InStr("nospace", " ") - 1)', "is -1"],
		['Chr$ past 255', 'Main = Chr$(Asc("z") + 200)', "is 322"],
		['an empty search string', 's = "abc"\n    Main = Left$(s, InStr(s, "") - 2)', "is -1"],
		['an empty string searched', 's = ""\n    Main = Left$("x", InStr(s, "a") - 1)', "is -1"],
		['a start after the match', 's = "abcabc"\n    Main = Left$(s, InStr(5, s, "a") - 5)', "is -5"],
		['a start past the end', 's = "abc"\n    Main = Left$(s, InStr(9, s, "a") - 1)', "is -1"],
		['binary compare by default', 's = "ABC"\n    Main = Left$(s, InStr(s, "b") - 1)', "is -1"],
		['InStrRev from the end', 's = "abca"\n    Main = Left$(s, InStrRev(s, "a") - 5)', "is -1"],
		['InStrRev with a start', 's = "abca"\n    Main = Left$(s, InStrRev(s, "a", 3) - 2)', "is -1"],
		['InStrRev of an empty search string', 's = "abc"\n    Main = Left$(s, InStrRev(s, "") - 4)', "is -1"],
		['InStrRev with a start past the end', 's = "abc"\n    Main = Left$(s, InStrRev(s, "a", 9) - 1)', "is -1"],
		['AscW', 'Main = Chr$(AscW("A") - 66)', "is -1"],
		['Len of a literal', 'Main = Space$(Len("ab") - 3)', "is -1"],
		['VBA.InStr', 's = "abc"\n    Main = Left$(s, VBA.InStr(s, "z") - 1)', "is -1"],
		['InStr inside InStr', 's = "aXbX"\n    Main = Left$(s, InStr(InStr(s, "X") + 1, s, "X") - 5)', "is -1"],
	];
	it.each(RAISES)('reports %s', (_name, body, text) => {
		const hits = argumentErrors(body);
		expect(hits).toHaveLength(1);
		expect(hits[0]).toContain(text);
	});

	const RUNS: ReadonlyArray<readonly [string, string, string?, string?]> = [
		['a space that is there', 's = "first last"\n    Main = Left$(s, InStr(s, " ") - 1)'],
		['one past a missing letter', 's = "abc"\n    Main = Mid$(s, InStr(s, "z") + 1)'],
		['Right$ of the rest', 's = "abc"\n    Main = Right$(s, Len(s) - InStr(s, "z"))'],
		['Mid$ just past the end', 's = "abc"\n    Main = Mid$(s, Len(s) + 1)'],
		['vbTextCompare', 's = "ABC"\n    Main = Left$(s, InStr(1, s, "b", vbTextCompare) - 1)'],
		['Option Compare Text', 's = "ABC"\n    Main = Left$(s, InStr(s, "b") - 1)', 'Option Compare Text'],
		// InStr(5, "abc", "") is 5: an empty string is found at the start, past the end too.
		['an empty search from past the end', 'Main = Left$("abcdef", InStr(5, "abc", "") - 1)'],
		// Not measured: what cannot be known stays quiet.
		['a string assigned twice', 's = "a b"\n    s = s & "c"\n    Main = Left$(s, InStr(s, " ") - 1)'],
		['Option Compare Database, where binary and text disagree', 's = "ABC"\n    Main = Left$(s, InStr(s, "b") - 1)', 'Option Compare Database'],
		['Asc above 127, which depends on the code page', 'Main = Chr$(Asc("\u00e9") + 200)'],
		['a project procedure named InStr', 's = "abc"\n    Main = Left$(s, InStr(s, "z") - 1)', '', 'Function InStr(a As String, b As String) As Long\n    InStr = 2\nEnd Function\n'],
	];
	it.each(RUNS)('stays quiet for %s', (_name, body, option, extra) => {
		expect(argumentErrors(body, option, extra)).toEqual([]);
	});
});
