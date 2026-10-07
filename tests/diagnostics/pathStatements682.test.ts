// Diagnostics tests: VBA file statements on paths the procedure deleted or
// made (issue #682). Measured on 2026-10-03 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Dim p As String, d As String, f As Integer\n    p = "C:\\t\\a.txt"\n    d = "C:\\t\\dir"\n    Main = "ok"\n    ${body}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.message);
}

const MAKE = 'f = FreeFile: Open p For Output As #f: Print #f, "abc": Close #f';
const GONE = 'On Error Resume Next: Kill p: On Error GoTo 0';

describe('file statements on a path the procedure deleted or made (issue #682)', () => {
	it.each([
		['Kill twice', `${MAKE}\n    Kill p\n    Kill p`, "Kill finds no file at p, which this procedure deleted or moved above. This will raise Run-time error '53': File not found."],
		['Open For Input after Kill', `${MAKE}\n    Kill p\n    f = FreeFile: Open p For Input As #f: Close #f`, "Run-time error '53'"],
		['Name onto a file it made', `${MAKE}\n    f = FreeFile: Open p & ".new" For Output As #f: Close #f\n    Name p As p & ".new"`, "Name finds p & \".new\" already there, a file this procedure made above. This will raise Run-time error '58': File already exists."],
		['MkDir twice', 'MkDir d\n    MkDir d', "Run-time error '75'"],
		['RmDir of a folder it filled', `MkDir d\n    f = FreeFile: Open d & "\\a.txt" For Output As #f: Close #f\n    RmDir d`, "RmDir removes only an empty folder. This will raise Run-time error '75'"],
		['RmDir of a folder holding a folder', 'MkDir d\n    MkDir d & "\\sub"\n    RmDir d', "Run-time error '75'"],
		['Kill after the Resume Next idiom', `${GONE}\n    Kill p`, "Run-time error '53'"],
		['FileLen after it', `${GONE}\n    Main = FileLen(p)`, "FileLen finds no file at p"],
		['GetAttr after it', `${GONE}\n    Main = GetAttr(p)`, "GetAttr finds no file at p"],
		['FileDateTime after it', `${GONE}\n    Main = FileDateTime(p)`, "FileDateTime finds no file at p"],
		['FileCopy after it', `${GONE}\n    FileCopy p, p & ".c"`, "FileCopy finds no file at p"],
		['Name after it', `${GONE}\n    Name p As p & ".n"`, "Name finds no file at p"],
		['Name after Kill', `${MAKE}\n    Kill p\n    Name p As p & ".x"`, "Name finds no file at p"],
		['RmDir after the idiom', 'On Error Resume Next: RmDir d: On Error GoTo 0\n    RmDir d', "RmDir finds no folder d, which this procedure removed above. This will raise Run-time error '76': Path not found."],
		['Kill under an error handler', `On Error GoTo Fail\n    ${MAKE}\n    Kill p\n    Kill p\n    Exit Function\nFail:\n    Main = Err.Number`, "Run-time error '53'"],
	])('checks %s under conservative runtime assumptions', (_label, body, message) => {
		const found = errors(body);
		if (/\bon\s+error\b/i.test(body)) {
			expect(found, body).toEqual([]);
			return;
		}
		expect(found, body).toHaveLength(1);
		expect(found[0], body).toContain(message);
	});

	it.each([
		['Append, which makes the file', `${GONE}\n    f = FreeFile: Open p For Append As #f: Close #f\n    Main = FileLen(p)`],
		['Dir of a deleted file', `${MAKE}\n    Kill p\n    Main = Len(Dir(p))`],
		['FileLen of a file just written', `${MAKE}\n    Main = FileLen(p)`],
		['Output, which makes it again', `${MAKE}\n    Kill p\n    f = FreeFile: Open p For Output As #f: Close #f\n    Kill p`],
		['Binary, which makes the file', `${GONE}\n    f = FreeFile: Open p For Binary As #f: Close #f\n    Main = FileLen(p)`],
		['Random, which makes the file', `${GONE}\n    f = FreeFile: Open p For Random As #f: Close #f\n    Main = FileLen(p)`],
		['MkDir and RmDir in turn', 'MkDir d\n    RmDir d\n    MkDir d\n    RmDir d'],
		['RmDir once the file in it is gone', `MkDir d\n    f = FreeFile: Open d & "\\a.txt" For Output As #f: Close #f\n    Kill d & "\\a.txt"\n    RmDir d`],
		['Kill of where Name moved it', `${MAKE}\n    Name p As p & ".new"\n    Kill p & ".new"`],
		['a second Kill under Resume Next', `${MAKE}\n    Kill p\n    On Error Resume Next\n    Kill p\n    On Error GoTo 0`],
		['a Kill a single-line If guards', `${MAKE}\n    Kill p\n    If Len(Dir(p)) > 0 Then Kill p`],
		['a Kill a block If guards', `${MAKE}\n    Kill p\n    If Dir(p) <> "" Then\n        Kill p\n    End If`],
		['a Kill an ElseIf guards', `${MAKE}\n    Kill p\n    If d = "" Then\n        Main = 1\n    ElseIf Dir(p) <> "" Then\n        Kill p\n    End If`],
		['a Kill a loop condition guards', `${MAKE}\n    Kill p\n    Do While Dir(p) <> ""\n        Kill p\n    Loop`],
		['MkDir twice under Resume Next', 'MkDir d\n    On Error Resume Next\n    MkDir d\n    On Error GoTo 0'],
		['Name once the target is gone', `${MAKE}\n    f = FreeFile: Open p & ".new" For Output As #f: Close #f\n    Kill p & ".new"\n    Name p As p & ".new"`],
		['FileCopy onto a file it made', `${MAKE}\n    f = FreeFile: Open p & ".new" For Output As #f: Close #f\n    FileCopy p, p & ".new"`],
		['a path that changed', `${MAKE}\n    Kill p\n    p = p & "2"\n    Kill p`],
		['a path a procedure may change', `${MAKE}\n    Kill p\n    Touch p\n    Kill p`],
		['FileCopy and Kill of the copy', `${MAKE}\n    FileCopy p, d & ".txt"\n    Kill p\n    Kill d & ".txt"`],
	])('stays quiet on %s', (_label, body) => {
		expect(errors(body), body).toEqual([]);
	});
});
