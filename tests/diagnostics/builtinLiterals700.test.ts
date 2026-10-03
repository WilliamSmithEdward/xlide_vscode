// Diagnostics tests: VBA built-ins whose failure the literals prove (issue
// #700). Measured on 2026-10-03 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Main = "ok"\n    ${body}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.message);
}

describe('built-ins a literal proves (issue #700)', () => {
	it.each([
		['Shell("")', 'Main = Shell("")', "Argument 'PathName' of 'Shell' is \"\"; this will raise Run-time error '5'"],
		['Environ(256)', 'Main = Len(Environ(256))', "Argument 'Expression' of 'Environ' is 256; this will raise Run-time error '5'"],
		['Environ(32767)', 'Main = Len(Environ(32767))', "Run-time error '5'"],
		['Environ(32768)', 'Main = Len(Environ(32768))', "Run-time error '6': Overflow"],
		['Environ("")', 'Main = Len(Environ(""))', "Argument 'Expression' of 'Environ' is \"\""],
		['Dir with attribute 64', 'Main = Len(Dir("C:\\", 64))', "Argument 'Attributes' of 'Dir' is 64"],
		['Dir with attribute 127', 'Main = Len(Dir("C:\\", 127))', "Run-time error '5'"],
		['Dir with attribute 256', 'Main = Len(Dir("C:\\", 256))', "Run-time error '5'"],
		['Dir with attribute -1', 'Main = Len(Dir("C:\\", -1))', "Run-time error '5'"],
		['SaveSetting with an empty AppName', 'SaveSetting "", "S", "k", "v"', "Argument 'AppName' of 'SaveSetting' is \"\""],
		['SaveSetting with an empty Section', 'SaveSetting "App", "", "k", "v"', "Argument 'Section' of 'SaveSetting' is \"\""],
		['SaveSetting with an empty Key', 'SaveSetting "App", "S", "", "v"', "Argument 'Key' of 'SaveSetting' is \"\""],
		['DeleteSetting with an empty AppName', 'DeleteSetting ""', "Argument 'AppName' of 'DeleteSetting' is \"\""],
		['GetSetting with an empty AppName', 'Main = GetSetting("", "S", "k", "d")', "Argument 'AppName' of 'GetSetting' is \"\""],
		['GetSetting with an empty Section', 'Main = GetSetting("App", "", "k", "d")', "Argument 'Section' of 'GetSetting' is \"\""],
	])('reports %s', (_label, body, message) => {
		const found = errors(body);
		expect(found, body).toHaveLength(1);
		expect(found[0], body).toContain(message);
	});

	it.each([
		'Main = Len(Environ(255))',
		'Main = Len(Environ("ZZ_NO_SUCH_VAR"))',
		'Main = Len(Dir("C:\\", 63))',
		'Main = Len(Dir("C:\\", 16))',
		'Main = Len(Dir("C:\\", vbDirectory))',
		'SaveSetting "App", "S", "k", ""',
		'Main = GetSetting("App", "S", "nope", "d")',
		'Main = IsEmpty(GetAllSettings("App", "Nope"))',
	])('stays quiet on %s', (body) => {
		expect(errors(body), body).toEqual([]);
	});
});

describe('DeleteSetting of what the procedure deleted (issue #700)', () => {
	const SAVE = 'SaveSetting "App", "S", "k", "v"';
	it.each([
		['a key twice', `${SAVE}\n    DeleteSetting "App", "S", "k"\n    DeleteSetting "App", "S", "k"`, 'DeleteSetting finds no such key: this procedure deleted it above.'],
		['a key of a deleted section', `${SAVE}\n    DeleteSetting "App", "S"\n    DeleteSetting "App", "S", "k"`, 'no such key'],
		['a section twice', `${SAVE}\n    DeleteSetting "App", "S"\n    DeleteSetting "App", "S"`, 'no such section'],
		['a key of a deleted application', `${SAVE}\n    DeleteSetting "App"\n    DeleteSetting "App", "S", "k"`, 'no such key'],
		['a key after the Resume Next idiom', 'On Error Resume Next: DeleteSetting "App": On Error GoTo 0\n    DeleteSetting "App", "S", "k"', 'no such key'],
	])('reports %s', (_label, body, message) => {
		const found = errors(body);
		expect(found, body).toHaveLength(1);
		expect(found[0], body).toContain(message);
		expect(found[0], body).toContain("Run-time error '5'");
	});

	it.each([
		['a key saved again', `${SAVE}\n    DeleteSetting "App", "S", "k"\n    ${SAVE}\n    DeleteSetting "App", "S", "k"`],
		['two keys', `${SAVE}\n    SaveSetting "App", "S", "k2", "v"\n    DeleteSetting "App", "S", "k"\n    DeleteSetting "App", "S", "k2"`],
		['a second delete under Resume Next', `${SAVE}\n    DeleteSetting "App", "S", "k"\n    On Error Resume Next\n    DeleteSetting "App", "S", "k"\n    On Error GoTo 0`],
		['a key never saved here', `${SAVE}\n    DeleteSetting "App", "S", "nope"`],
		['another application', `${SAVE}\n    DeleteSetting "App"\n    DeleteSetting "Other", "S", "k"`],
		['a delete a block may repeat', `${SAVE}\n    DeleteSetting "App", "S", "k"\n    If Len(GetSetting("App", "S", "k")) > 0 Then\n        DeleteSetting "App", "S", "k"\n    End If`],
		['a call between', `${SAVE}\n    DeleteSetting "App", "S", "k"\n    Restore\n    DeleteSetting "App", "S", "k"`],
	])('stays quiet on %s', (_label, body) => {
		expect(errors(body), body).toEqual([]);
	});
});
